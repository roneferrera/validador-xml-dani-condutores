import streamlit as st
import pandas as pd
import io
import re
import zipfile
from lxml import etree

# ─────────────────────────────────────────────
# TEMA THOMSON REUTERS / DOMÍNIO SISTEMAS
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Enriquecedor de NF-e — DNI",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    .stApp { background-color: #F5F5F5; }
    .main-header {
        background: linear-gradient(135deg, #C8102E 0%, #8B0000 100%);
        padding: 20px 32px; border-radius: 8px; margin-bottom: 20px;
    }
    .main-header h1 { color: white !important; font-size: 1.5rem !important;
        font-weight: 700 !important; margin: 0 !important; }
    .main-header p { color: rgba(255,255,255,0.85) !important;
        font-size: 0.82rem !important; margin: 4px 0 0 0 !important; }
    .section-card { background: white; border: 1px solid #E0E0E0;
        border-left: 4px solid #C8102E; border-radius: 6px;
        padding: 18px 22px; margin-bottom: 14px; }
    .section-title { color: #C8102E; font-size: 0.85rem; font-weight: 700;
        text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 10px; }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #C8102E, #8B0000) !important;
        color: white !important; border: none !important; border-radius: 6px !important;
        font-weight: 600 !important; font-size: 1rem !important;
        padding: 12px 32px !important; width: 100%;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #A00D25, #6B0000) !important;
        box-shadow: 0 4px 12px rgba(200,16,46,0.3) !important;
    }
    [data-testid="metric-container"] { background: white;
        border: 1px solid #E0E0E0; border-radius: 6px; padding: 10px 14px; }
    .footer { text-align: center; color: #999; font-size: 0.75rem;
        margin-top: 28px; padding-top: 14px; border-top: 1px solid #E0E0E0; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>⚡ Enriquecedor de NF-e — DNI</h1>
    <p>Thomson Reuters · Domínio Sistemas · Processamento Fiscal Automatizado</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def limpar_chave(valor: str) -> str:
    if not valor:
        return ""
    return re.sub(r"[^0-9]", "", str(valor).strip())

def limpar_valor(valor) -> float:
    if valor is None:
        return 0.0
    s = re.sub(r"[^\d\.\-]", "", str(valor).strip().replace(",", "."))
    try:
        return float(s)
    except ValueError:
        return 0.0

def fmt(v: float) -> str:
    return f"{v:.2f}"

def safe_int_cst(val: str) -> str:
    try:
        return str(int(float(val))).zfill(2)
    except Exception:
        return str(val).zfill(2)

# ─────────────────────────────────────────────
# LEITURA DO XLSX
# ─────────────────────────────────────────────
def ler_xlsx(conteudo_bytes: bytes):
    indexado = {}
    try:
        df = pd.read_excel(io.BytesIO(conteudo_bytes), dtype=str, engine="openpyxl")
    except Exception as e:
        st.error(f"Erro ao ler XLSX: {e}")
        return indexado

    df.columns = [str(c).strip() for c in df.columns]

    for _, row in df.iterrows():
        reg = {k: (str(v).strip() if pd.notna(v) else "") for k, v in row.items()}

        chave_nfe = limpar_chave(reg.get("Chave Nfe/Cte", ""))
        seq_raw   = reg.get("Sequencia", "0")
        try:
            seq = int(float(seq_raw)) if seq_raw else 0
        except Exception:
            seq = 0

        if len(chave_nfe) < 44 or seq == 0:
            continue

        vlr_documento = limpar_valor(reg.get("Vlr Documento", "0"))
        vlr_icms      = limpar_valor(reg.get("Vlr Icms", "0"))
        vlr_icms_st   = limpar_valor(reg.get("Vlr Icms St", "0"))
        vlr_ipi       = limpar_valor(reg.get("Vlr Ipi", "0"))
        bc_pis_cofins = max(0.0, vlr_documento - vlr_icms - vlr_icms_st - vlr_ipi)

        indexado[(chave_nfe, seq)] = {
            "cfop":          reg.get("Cfop", "").strip(),
            "cod_item":      reg.get("Cod Item", "").strip(),
            "desc_item":     reg.get("Desc Item", "").strip(),
            "ncm":           reg.get("NCM", "").strip(),
            "nro_documento": reg.get("Nro Documento", "").strip(),
            "razao_social":  reg.get("Razao Social", "").strip(),
            "cst_icms":      safe_int_cst(reg.get("CST ICMS", "0")),
            "base_icms":     limpar_valor(reg.get("Base Icms", "0")),
            "perc_icms":     limpar_valor(reg.get("Perc ICms", "0")),
            "vlr_icms":      vlr_icms,
            "base_icms_st":  limpar_valor(reg.get("Base Icms St", "0")),
            "vlr_icms_st":   vlr_icms_st,
            "cst_ipi":       safe_int_cst(reg.get("CST IPI", "50")),
            "base_ipi":      limpar_valor(reg.get("Base Ipi", "0")),
            "perc_ipi":      limpar_valor(reg.get("Perc Ipi", "0")),
            "vlr_ipi":       vlr_ipi,
            "cst_pis":       safe_int_cst(reg.get("CST PIS", "1")),
            "perc_pis":      limpar_valor(reg.get("Perc Pis", "0")),
            "vlr_pis":       limpar_valor(reg.get("Vlr Pis", "0")),
            "cst_cofins":    safe_int_cst(reg.get("CST COFINS", "1")),
            "perc_cofins":   limpar_valor(reg.get("Perc Cofins", "0")),
            "vlr_cofins":    limpar_valor(reg.get("Vlr Cofins", "0")),
            "bc_pis_cofins": bc_pis_cofins,
            "vlr_documento": vlr_documento,
        }

    return indexado

# ─────────────────────────────────────────────
# HELPERS XML
# ─────────────────────────────────────────────
NS = "http://www.portalfiscal.inf.br/nfe"

def tag(nome): return f"{{{NS}}}{nome}"

def local(elem):
    return elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

def find(elem, *nomes):
    atual = elem
    for nome in nomes:
        enc = None
        for filho in atual:
            if local(filho) == nome:
                enc = filho
                break
        if enc is None:
            return None
        atual = enc
    return atual

def set_elem(pai, nome_local, valor, idx_apos=None):
    """Define texto de filho, criando se necessário e posicionando após idx_apos."""
    for filho in pai:
        if local(filho) == nome_local:
            filho.text = valor
            return filho
    novo = etree.SubElement(pai, tag(nome_local))
    novo.text = valor
    if idx_apos is not None:
        pai.remove(novo)
        pai.insert(idx_apos + 1, novo)
    return novo

# ─────────────────────────────────────────────
# PROCESSAMENTO ICMS ST (ORDEM SEFAZ)
# ─────────────────────────────────────────────
CST_COM_ST = {"10", "30", "70", "90"}

def aplicar_icms(icms_filho, dados):
    """Aplica ICMS próprio + ST conforme ordem do schema SEFAZ."""
    cst = dados["cst_icms"]
    modificado = False

    # CST
    el = find(icms_filho, "CST")
    if el is not None:
        el.text = cst
        modificado = True

    # ICMS próprio
    for tag_nome, val in [
        ("vBC",   fmt(dados["base_icms"])),
        ("pICMS", fmt(dados["perc_icms"])),
        ("vICMS", fmt(dados["vlr_icms"])),
    ]:
        el = find(icms_filho, tag_nome)
        if el is not None:
            el.text = val
            modificado = True

    # ST — só aplica se CST com ST e houver valores
    if cst in CST_COM_ST and (dados["base_icms_st"] > 0 or dados["vlr_icms_st"] > 0):
        # modBCST — mantém existente ou usa 4 (MVA)
        el_mod = find(icms_filho, "modBCST")
        if el_mod is None:
            # Insere após vICMS
            idx = next((i for i, f in enumerate(icms_filho) if local(f) == "vICMS"), None)
            el_mod = set_elem(icms_filho, "modBCST", "4", idx)
        mod_val = (el_mod.text or "4").strip()

        # pMVAST — obrigatório se modBCST == 4
        if mod_val == "4":
            el_pmva = find(icms_filho, "pMVAST")
            if el_pmva is None:
                idx_mod = next((i for i, f in enumerate(icms_filho) if local(f) == "modBCST"), None)
                el_pmva = set_elem(icms_filho, "pMVAST", "0.00", idx_mod)
            # mantém valor original se já existir

        # vBCST / pICMSST / vICMSST
        for tag_nome, val in [
            ("vBCST",    fmt(dados["base_icms_st"])),
            ("pICMSST",  fmt(dados["perc_icms"])),
            ("vICMSST",  fmt(max(0.0, dados["vlr_icms_st"]))),
        ]:
            el = find(icms_filho, tag_nome)
            if el is not None:
                el.text = val
                modificado = True

    return modificado


# ─────────────────────────────────────────────
# PROCESSAMENTO XML PRINCIPAL
# ─────────────────────────────────────────────
def processar_xml(conteudo_xml, nome_arquivo, dados_indexados, cfops_ativas):
    try:
        tree = etree.fromstring(conteudo_xml)
    except Exception as e:
        return None, f"XML inválido: {e}", "erro", []

    inf_nfe = find(tree, "NFe", "infNFe") or find(tree, "infNFe")
    if inf_nfe is None:
        return None, "infNFe não encontrado", "erro", []

    chave_xml = re.sub(r"[^0-9]", "", inf_nfe.get("Id", ""))
    if len(chave_xml) < 44:
        return None, f"Chave inválida: {chave_xml}", "erro", []

    det_elements = [f for f in inf_nfe if local(f) == "det"]

    itens_validos = []
    for det in det_elements:
        try:
            n_item = int(det.get("nItem", "0"))
        except Exception:
            continue
        dados = dados_indexados.get((chave_xml, n_item))
        if dados is None:
            continue
        if dados.get("cfop", "") in cfops_ativas:
            itens_validos.append((det, n_item, dados))

    if not itens_validos:
        return None, "nenhum item com CFOP válido — não alterado", "info", []

    modificado = False
    diferencas = []  # para excel de conferência

    for det, n_item, dados in itens_validos:
        prod    = find(det, "prod")
        imposto = find(det, "imposto")
        if prod is None or imposto is None:
            continue

        # ── Coleta valores ANTES (para conferência) ──
        def _get(elem, *path):
            el = find(elem, *path)
            return el.text.strip() if el is not None and el.text else "0"

        antes = {
            "vBC_icms":   _get(imposto, "ICMS", *["ICMS" + dados["cst_icms"].zfill(2)], "vBC") if False else "0",
            "vICMS":      "0",
            "vBCST":      "0",
            "vICMSST":    "0",
            "vIPI":       "0",
            "vBC_pis":    "0",
            "vPIS":       "0",
            "vBC_cofins": "0",
            "vCOFINS":    "0",
        }

        # Coleta real dos valores antes
        icms_pai = find(imposto, "ICMS")
        if icms_pai:
            for icms_f in icms_pai:
                if local(icms_f).startswith("ICMS"):
                    antes["vBC_icms"] = _get(icms_f, "vBC") or "0"
                    antes["vICMS"]    = _get(icms_f, "vICMS") or "0"
                    antes["vBCST"]    = _get(icms_f, "vBCST") or "0"
                    antes["vICMSST"]  = _get(icms_f, "vICMSST") or "0"
                    break
        ipi_el = find(imposto, "IPI", "IPITrib")
        if ipi_el:
            antes["vIPI"] = _get(ipi_el, "vIPI") or "0"
        pis_pai = find(imposto, "PIS")
        if pis_pai:
            for pf in pis_pai:
                antes["vBC_pis"] = _get(pf, "vBC") or "0"
                antes["vPIS"]    = _get(pf, "vPIS") or "0"
                break
        cof_pai = find(imposto, "COFINS")
        if cof_pai:
            for cf in cof_pai:
                antes["vBC_cofins"] = _get(cf, "vBC") or "0"
                antes["vCOFINS"]    = _get(cf, "vCOFINS") or "0"
                break

        # ── CFOP ──
        el = find(prod, "CFOP")
        if el is not None and dados["cfop"]:
            el.text = dados["cfop"]
            modificado = True

        # ── NCM ──
        el = find(prod, "NCM")
        if el is not None and dados["ncm"]:
            el.text = dados["ncm"]
            modificado = True

        # ── ICMS ──
        if icms_pai is not None:
            for icms_f in icms_pai:
                if local(icms_f).startswith("ICMS"):
                    if aplicar_icms(icms_f, dados):
                        modificado = True
                    break

        # ── IPI ──
        ipi_elem = find(imposto, "IPI")
        if ipi_elem is not None:
            ipi_trib = find(ipi_elem, "IPITrib")
            if ipi_trib is not None:
                for tag_nome, val in [
                    ("CST",  dados["cst_ipi"]),
                    ("vBC",  fmt(dados["base_ipi"])),
                    ("pIPI", fmt(dados["perc_ipi"])),
                    ("vIPI", fmt(dados["vlr_ipi"])),
                ]:
                    el = find(ipi_trib, tag_nome)
                    if el is not None:
                        el.text = val
                        modificado = True

        # ── PIS ──
        pis_pai = find(imposto, "PIS")
        if pis_pai is not None:
            for pf in pis_pai:
                for tag_nome, val in [
                    ("CST",  dados["cst_pis"]),
                    ("vBC",  fmt(dados["bc_pis_cofins"])),
                    ("pPIS", fmt(dados["perc_pis"])),
                    ("vPIS", fmt(dados["vlr_pis"])),
                ]:
                    el = find(pf, tag_nome)
                    if el is not None:
                        el.text = val
                        modificado = True
                break

        # ── COFINS ──
        cof_pai = find(imposto, "COFINS")
        if cof_pai is not None:
            for cf in cof_pai:
                for tag_nome, val in [
                    ("CST",     dados["cst_cofins"]),
                    ("vBC",     fmt(dados["bc_pis_cofins"])),
                    ("pCOFINS", fmt(dados["perc_cofins"])),
                    ("vCOFINS", fmt(dados["vlr_cofins"])),
                ]:
                    el = find(cf, tag_nome)
                    if el is not None:
                        el.text = val
                        modificado = True
                break

        # ── Coleta valores DEPOIS ──
        depois = {
            "vBC_icms":   fmt(dados["base_icms"]),
            "vICMS":      fmt(dados["vlr_icms"]),
            "vBCST":      fmt(dados["base_icms_st"]),
            "vICMSST":    fmt(dados["vlr_icms_st"]),
            "vIPI":       fmt(dados["vlr_ipi"]),
            "vBC_pis":    fmt(dados["bc_pis_cofins"]),
            "vPIS":       fmt(dados["vlr_pis"]),
            "vBC_cofins": fmt(dados["bc_pis_cofins"]),
            "vCOFINS":    fmt(dados["vlr_cofins"]),
        }

        # Detecta diferenças
        campos_conf = [
            ("vBC ICMS",   "vBC_icms"),
            ("vICMS",      "vICMS"),
            ("vBCST",      "vBCST"),
            ("vICMSST",    "vICMSST"),
            ("vIPI",       "vIPI"),
            ("BC PIS",     "vBC_pis"),
            ("vPIS",       "vPIS"),
            ("BC COFINS",  "vBC_cofins"),
            ("vCOFINS",    "vCOFINS"),
        ]
        tem_diff = any(
            round(float(antes[k]), 2) != round(float(depois[k]), 2)
            for _, k in campos_conf
        )

        if tem_diff:
            xprod_el = find(prod, "xProd")
            xprod = xprod_el.text if xprod_el is not None else ""
            row_conf = {
                "Arquivo":      nome_arquivo,
                "Chave":        chave_xml,
                "nItem":        n_item,
                "Cod Item":     dados["cod_item"],
                "Desc Item":    xprod,
                "CFOP":         dados["cfop"],
            }
            for label, k in campos_conf:
                row_conf[f"{label} (antes)"] = antes[k]
                row_conf[f"{label} (depois)"] = depois[k]
                try:
                    diff = round(float(depois[k]) - float(antes[k]), 2)
                except Exception:
                    diff = ""
                row_conf[f"{label} (diff)"] = diff
            diferencas.append(row_conf)

    if not modificado:
        return None, "nenhuma alteração aplicada", "info", []

    recalcular_totais(inf_nfe)
    xml_out = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", pretty_print=False)
    return xml_out, f"alterado com sucesso ({len(itens_validos)} itens)", "ok", diferencas


# ─────────────────────────────────────────────
# RECALCULA ICMSTot
# ─────────────────────────────────────────────
def recalcular_totais(inf_nfe):
    totais = {k: 0.0 for k in [
        "vBC","vICMS","vBCST","vST","vIPI","vIPIDevol",
        "vPIS","vCOFINS","vProd"
    ]}

    for filho in inf_nfe:
        if local(filho) != "det":
            continue
        imposto = find(filho, "imposto")
        prod    = find(filho, "prod")

        if prod:
            el = find(prod, "vProd")
            if el is not None and el.text:
                try: totais["vProd"] += float(el.text)
                except: pass

        if imposto:
            icms_pai = find(imposto, "ICMS")
            if icms_pai:
                for icms_f in icms_pai:
                    if local(icms_f).startswith("ICMS"):
                        for k, t in [("vBC","vBC"),("vICMS","vICMS"),
                                     ("vBCST","vBCST"),("vICMSST","vST")]:
                            el = find(icms_f, k)
                            if el is not None and el.text:
                                try: totais[t] += float(el.text)
                                except: pass
                        break

            ipi_trib = find(imposto, "IPI", "IPITrib")
            if ipi_trib:
                el = find(ipi_trib, "vIPI")
                if el is not None and el.text:
                    try: totais["vIPI"] += float(el.text)
                    except: pass

            pis_pai = find(imposto, "PIS")
            if pis_pai:
                for pf in pis_pai:
                    el = find(pf, "vPIS")
                    if el is not None and el.text:
                        try: totais["vPIS"] += float(el.text)
                        except: pass
                    break

            cof_pai = find(imposto, "COFINS")
            if cof_pai:
                for cf in cof_pai:
                    el = find(cf, "vCOFINS")
                    if el is not None and el.text:
                        try: totais["vCOFINS"] += float(el.text)
                        except: pass
                    break

        imp_devol = find(filho, "impostoDevol")
        if imp_devol:
            el = find(imp_devol, "IPI", "vIPIDevol")
            if el is not None and el.text:
                try: totais["vIPIDevol"] += float(el.text)
                except: pass

    icms_tot = find(inf_nfe, "total", "ICMSTot")
    if icms_tot is None:
        return

    mapa = {
        "vBC":      fmt(totais["vBC"]),
        "vICMS":    fmt(totais["vICMS"]),
        "vBCST":    fmt(totais["vBCST"]),
        "vST":      fmt(totais["vST"]),
        "vIPI":     fmt(totais["vIPI"]),
        "vIPIDevol":fmt(totais["vIPIDevol"]),
        "vPIS":     fmt(totais["vPIS"]),
        "vCOFINS":  fmt(totais["vCOFINS"]),
        "vProd":    fmt(totais["vProd"]),
        "vNF":      fmt(totais["vProd"] + totais["vST"] + totais["vIPI"]),
    }
    for tag_nome, val in mapa.items():
        el = find(icms_tot, tag_nome)
        if el is not None:
            el.text = val


# ─────────────────────────────────────────────
# GERA EXCEL DE CONFERÊNCIA
# ─────────────────────────────────────────────
def gerar_excel_conferencia(todas_diferencas: list) -> bytes:
    if not todas_diferencas:
        return b""
    df = pd.DataFrame(todas_diferencas)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Diferenças")
        ws = writer.sheets["Diferenças"]

        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        # Cabeçalho vermelho
        header_fill = PatternFill("solid", fgColor="C8102E")
        header_font = Font(color="FFFFFF", bold=True, size=10)
        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        # Destaca células com diff != 0
        diff_fill_pos = PatternFill("solid", fgColor="C6EFCE")  # verde
        diff_fill_neg = PatternFill("solid", fgColor="FFC7CE")  # vermelho
        diff_font_pos = Font(color="276221", bold=True, size=9)
        diff_font_neg = Font(color="9C0006", bold=True, size=9)

        cols_diff = [i+1 for i, c in enumerate(df.columns) if "(diff)" in c]

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(size=9)

            for col_idx in cols_diff:
                cell = row[col_idx - 1]
                try:
                    v = float(cell.value or 0)
                    if v > 0:
                        cell.fill = diff_fill_pos
                        cell.font = diff_font_pos
                    elif v < 0:
                        cell.fill = diff_fill_neg
                        cell.font = diff_font_neg
                except Exception:
                    pass

        # Ajusta largura das colunas
        for col_idx, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
            max_len = max(
                len(str(col[0].value or "")),
                max((len(str(ws.cell(r, col_idx).value or "")) for r in range(2, ws.max_row + 1)), default=0)
            )
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 30)

        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "A2"

    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# INTERFACE
# ─────────────────────────────────────────────

# 1 — XLSX
st.markdown('<div class="section-card"><div class="section-title">📂 1. Arquivo de Entrada (XLSX — Domínio Sistemas)</div>', unsafe_allow_html=True)
arquivo_xlsx = st.file_uploader("Selecione o arquivo .xlsx exportado do Domínio", type=["xlsx"], key="xlsx")
st.markdown('</div>', unsafe_allow_html=True)

# 2 — XMLs
st.markdown('<div class="section-card"><div class="section-title">📄 2. Arquivos XML de NF-e</div>', unsafe_allow_html=True)
arquivos_xml = st.file_uploader(
    "Selecione um ou mais XMLs ou um arquivo ZIP contendo XMLs",
    type=["xml", "zip"],
    accept_multiple_files=True,
    key="xmls"
)
st.markdown('</div>', unsafe_allow_html=True)

# 3 — CFOPs
st.markdown('<div class="section-card"><div class="section-title">🔢 3. CFOPs a Processar</div>', unsafe_allow_html=True)

CFOPS_DEFAULT = ["1201","1202","1410","1411","2201","2202","2410","2411","1949","2949","2603"]

if "cfops_lista" not in st.session_state:
    st.session_state["cfops_lista"] = list(CFOPS_DEFAULT)

col_inp, col_add = st.columns([3, 1])
with col_inp:
    nova_cfop = st.text_input(
        "Adicionar CFOP (4 dígitos):",
        max_chars=4,
        placeholder="ex: 1410",
        key="nova_cfop_input",
        label_visibility="collapsed",
    )
with col_add:
    if st.button("➕ Adicionar", key="btn_add_cfop"):
        cfop_clean = re.sub(r"[^0-9]", "", nova_cfop.strip())
        if len(cfop_clean) == 4 and cfop_clean not in st.session_state["cfops_lista"]:
            st.session_state["cfops_lista"].append(cfop_clean)
            st.rerun()
        elif len(cfop_clean) != 4:
            st.warning("CFOP deve ter exatamente 4 dígitos.")
        else:
            st.info(f"CFOP {cfop_clean} já está na lista.")

# Exibe CFOPs com botão de remover individual
cfops_para_remover = []
if st.session_state["cfops_lista"]:
    cols_cfop = st.columns(min(len(st.session_state["cfops_lista"]), 8))
    for i, cfop in enumerate(st.session_state["cfops_lista"]):
        with cols_cfop[i % 8]:
            if st.button(f"❌ {cfop}", key=f"rm_{cfop}_{i}", help=f"Remover CFOP {cfop}"):
                cfops_para_remover.append(cfop)

if cfops_para_remover:
    for c in cfops_para_remover:
        if c in st.session_state["cfops_lista"]:
            st.session_state["cfops_lista"].remove(c)
    st.rerun()

# Remover em lote
col_lote, col_reset = st.columns([3, 1])
with col_lote:
    remover_lote = st.text_input(
        "Remover CFOPs em lote (separadas por vírgula):",
        placeholder="ex: 1949, 2949, 2603",
        key="remover_lote_input",
        label_visibility="collapsed",
    )
with col_reset:
    if st.button("🔄 Restaurar padrão", key="btn_reset_cfop"):
        st.session_state["cfops_lista"] = list(CFOPS_DEFAULT)
        st.rerun()

if st.button("🗑️ Remover em lote", key="btn_remover_lote"):
    cfops_lote = [re.sub(r"[^0-9]", "", c.strip()) for c in remover_lote.split(",") if c.strip()]
    for c in cfops_lote:
        if c in st.session_state["cfops_lista"]:
            st.session_state["cfops_lista"].remove(c)
    st.rerun()

st.caption(f"CFOPs ativas: **{', '.join(sorted(st.session_state['cfops_lista'])) or 'Nenhuma'}**")
st.markdown('</div>', unsafe_allow_html=True)

# 4 — Processar
st.markdown('<div class="section-card"><div class="section-title">▶️ 4. Processar</div>', unsafe_allow_html=True)

if st.button("▶️ Processar XMLs", type="primary", key="btn_processar"):
    if not arquivo_xlsx:
        st.error("❌ Selecione o arquivo XLSX.")
        st.stop()
    if not arquivos_xml:
        st.error("❌ Selecione ao menos um XML ou ZIP.")
        st.stop()
    if not st.session_state["cfops_lista"]:
        st.error("❌ Adicione ao menos uma CFOP.")
        st.stop()

    cfops_ativas = set(st.session_state["cfops_lista"])

    with st.spinner("🔄 Lendo XLSX..."):
        dados_indexados = ler_xlsx(arquivo_xlsx.read())

    if not dados_indexados:
        st.error("❌ Nenhum item válido encontrado no XLSX.")
        st.stop()

    col_a, col_b, col_c = st.columns(3)
    col_a.metric("📊 Itens indexados", len(dados_indexados))
    col_b.metric("🔢 CFOPs ativas", len(cfops_ativas))
    col_c.metric("📁 Arquivos recebidos", len(arquivos_xml))

    with st.expander("🔍 Debug — primeiros 5 itens indexados"):
        for i, ((ch, seq), d) in enumerate(list(dados_indexados.items())[:5]):
            st.code(
                f"Chave: {ch} | Seq: {seq} | CFOP: {d['cfop']}\n"
                f"Item: {d['cod_item']} | vICMS: {d['vlr_icms']} | "
                f"vICMSST: {d['vlr_icms_st']} | vIPI: {d['vlr_ipi']} | "
                f"BC PIS/COFINS: {d['bc_pis_cofins']}"
            )

    # Coleta XMLs (suporta ZIP)
    xmls_para_processar = {}
    for arq in arquivos_xml:
        if arq.name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(arq.read())) as zf:
                for nome in zf.namelist():
                    if nome.lower().endswith(".xml"):
                        xmls_para_processar[nome] = zf.read(nome)
        else:
            xmls_para_processar[arq.name] = arq.read()

    resultados = []
    xmls_modificados = {}
    todas_diferencas = []
    progress = st.progress(0)
    total = len(xmls_para_processar)

    for idx, (nome_arq, conteudo) in enumerate(xmls_para_processar.items()):
        xml_out, msg, status, diffs = processar_xml(
            conteudo, nome_arq, dados_indexados, cfops_ativas
        )
        resultados.append((nome_arq, msg, status))
        if xml_out:
            xmls_modificados[nome_arq] = xml_out
        todas_diferencas.extend(diffs)
        progress.progress((idx + 1) / total)

    progress.empty()

    # Resumo
    ok_c   = sum(1 for _, _, s in resultados if s == "ok")
    info_c = sum(1 for _, _, s in resultados if s == "info")
    err_c  = sum(1 for _, _, s in resultados if s == "erro")

    col_x, col_y, col_z, col_w = st.columns(4)
    col_x.metric("✅ Alterados",     ok_c)
    col_y.metric("ℹ️ Sem alteração", info_c)
    col_z.metric("❌ Erros",         err_c)
    col_w.metric("📋 Itens c/ diff", len(todas_diferencas))

    st.subheader("Resultados por Arquivo")
    for nome, msg, status in resultados:
        if status == "ok":
            st.success(f"✅ **{nome}**: {msg}")
        elif status == "info":
            st.info(f"ℹ️ **{nome}**: {msg}")
        else:
            st.error(f"❌ **{nome}**: {msg}")

    # Download XMLs
    if xmls_modificados:
        st.subheader("⬇️ Download XMLs")
        if len(xmls_modificados) == 1:
            nome_arq, conteudo_arq = list(xmls_modificados.items())[0]
            st.download_button(
                label=f"⬇️ Baixar {nome_arq}",
                data=conteudo_arq,
                file_name=nome_arq,
                mime="application/xml",
            )
        else:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for nome_arq, conteudo_arq in xmls_modificados.items():
                    zf.writestr(nome_arq, conteudo_arq)
            buf.seek(0)
            st.download_button(
                label=f"⬇️ Baixar todos ({len(xmls_modificados)} XMLs) como ZIP",
                data=buf,
                file_name="xmls_modificados_dni.zip",
                mime="application/zip",
            )

    # Download Excel de Conferência
    if todas_diferencas:
        st.subheader("📊 Excel de Conferência")
        excel_bytes = gerar_excel_conferencia(todas_diferencas)
        st.download_button(
            label=f"📥 Baixar Excel de Conferência ({len(todas_diferencas)} itens com diferença)",
            data=excel_bytes,
            file_name="conferencia_diferencas_dni.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        if xmls_modificados:
            st.info("✅ Nenhuma diferença detectada entre os valores originais e os do XLSX.")

    if not xmls_modificados:
        st.warning("⚠️ Nenhum XML foi modificado.")

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("""
<div class="footer">
    Thomson Reuters · Domínio Sistemas · Enriquecedor NF-e v3.0 · DNI
</div>
""", unsafe_allow_html=True)
