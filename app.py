import streamlit as st
import pandas as pd
import io
import re
import zipfile
from lxml import etree

# ─────────────────────────────────────────────
# TEMA THOMSON REUTERS — igual ao RPA TXT
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Enriquecedor de NF-e — DNI",
    page_icon="🟠",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: 'Segoe UI', 'Arial', sans-serif;
    color: #444444;
}
h1, h2, h3 { color: #FF8000; font-weight: 700; }

.main-header {
    background: #444444;
    padding: 22px 28px 16px 28px;
    border-radius: 8px;
    border-top: 6px solid #FF8000;
    margin-bottom: 24px;
}
.main-header h2 {
    color: #FF8000 !important;
    margin: 0;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 1.5rem;
}
.main-header p {
    color: #DDDDDD !important;
    margin: 6px 0 0 0;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 0.85rem;
}

.section-card {
    background: white;
    border: 1px solid #E0E0E0;
    border-left: 4px solid #FF8000;
    border-radius: 6px;
    padding: 18px 22px;
    margin-bottom: 14px;
}
.section-title {
    color: #FF8000;
    font-size: 0.85rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 10px;
}

.stButton > button {
    background-color: #FF8000 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 4px !important;
    font-weight: bold !important;
}
.stButton > button:hover {
    background-color: #D64001 !important;
    color: #FFFFFF !important;
}
.stDownloadButton > button {
    background-color: #FF8000 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 4px !important;
    font-weight: bold !important;
}
.stDownloadButton > button:hover {
    background-color: #D64001 !important;
    color: #FFFFFF !important;
}
[data-testid="metric-container"] {
    background-color: #E9E9E9;
    border-left: 4px solid #FF8000;
    border-radius: 4px;
    padding: 10px;
}
hr { border-color: #FF8000; }
.footer {
    text-align: center;
    color: #999;
    font-size: 0.75rem;
    margin-top: 28px;
    padding-top: 14px;
    border-top: 1px solid #E0E0E0;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h2>🧾 Enriquecedor de NF-e — DNI</h2>
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
        return str(val).zfill(2) if val else "00"

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
            "cfop":           reg.get("Cfop", "").strip(),
            "cod_item":       reg.get("Cod Item", "").strip(),
            "desc_item":      reg.get("Desc Item", "").strip(),
            "ncm":            reg.get("NCM", "").strip(),
            "nro_documento":  reg.get("Nro Documento", "").strip(),
            "razao_social":   reg.get("Razao Social", "").strip(),
            "cnpj":           reg.get("CNPJ-CPF", "").strip(),
            "data_entrada":   reg.get("Data Entrada", "").strip(),
            "data_emissao":   reg.get("Data Emissao", "").strip(),
            "uf":             reg.get("UF", "").strip(),
            "municipio":      reg.get("Municipio", "").strip(),
            "cst_icms":       safe_int_cst(reg.get("CST ICMS", "0")),
            "base_icms":      limpar_valor(reg.get("Base Icms", "0")),
            "perc_icms":      limpar_valor(reg.get("Perc ICms", "0")),
            "vlr_icms":       vlr_icms,
            "base_icms_st":   limpar_valor(reg.get("Base Icms St", "0")),
            "vlr_icms_st":    vlr_icms_st,
            "cst_ipi":        safe_int_cst(reg.get("CST IPI", "50")),
            "base_ipi":       limpar_valor(reg.get("Base Ipi", "0")),
            "perc_ipi":       limpar_valor(reg.get("Perc Ipi", "0")),
            "vlr_ipi":        vlr_ipi,
            "cst_pis":        safe_int_cst(reg.get("CST PIS", "1")),
            "perc_pis":       limpar_valor(reg.get("Perc Pis", "0")),
            "vlr_pis":        limpar_valor(reg.get("Vlr Pis", "0")),
            "cst_cofins":     safe_int_cst(reg.get("CST COFINS", "1")),
            "perc_cofins":    limpar_valor(reg.get("Perc Cofins", "0")),
            "vlr_cofins":     limpar_valor(reg.get("Vlr Cofins", "0")),
            "bc_pis_cofins":  bc_pis_cofins,
            "vlr_documento":  vlr_documento,
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

# ─────────────────────────────────────────────
# ICMS ST
# ─────────────────────────────────────────────
CST_COM_ST = {"10", "30", "70", "90"}

def aplicar_icms(icms_filho, dados):
    cst = dados["cst_icms"]
    modificado = False

    el = find(icms_filho, "CST")
    if el is not None:
        el.text = cst
        modificado = True

    for tag_nome, val in [
        ("vBC",   fmt(dados["base_icms"])),
        ("pICMS", fmt(dados["perc_icms"])),
        ("vICMS", fmt(dados["vlr_icms"])),
    ]:
        el = find(icms_filho, tag_nome)
        if el is not None:
            el.text = val
            modificado = True

    if cst in CST_COM_ST and (dados["base_icms_st"] > 0 or dados["vlr_icms_st"] > 0):
        el_mod = find(icms_filho, "modBCST")
        if el_mod is None:
            idx = next((i for i, f in enumerate(icms_filho) if local(f) == "vICMS"), None)
            el_mod = etree.SubElement(icms_filho, tag("modBCST"))
            el_mod.text = "4"
            if idx is not None:
                icms_filho.remove(el_mod)
                icms_filho.insert(idx + 1, el_mod)

        mod_val = (el_mod.text or "4").strip()
        if mod_val == "4":
            el_pmva = find(icms_filho, "pMVAST")
            if el_pmva is None:
                idx_mod = next((i for i, f in enumerate(icms_filho) if local(f) == "modBCST"), None)
                el_pmva = etree.SubElement(icms_filho, tag("pMVAST"))
                el_pmva.text = "0.00"
                if idx_mod is not None:
                    icms_filho.remove(el_pmva)
                    icms_filho.insert(idx_mod + 1, el_pmva)

        for tag_nome, val in [
            ("vBCST",   fmt(dados["base_icms_st"])),
            ("pICMSST", fmt(dados["perc_icms"])),
            ("vICMSST", fmt(max(0.0, dados["vlr_icms_st"]))),
        ]:
            el = find(icms_filho, tag_nome)
            if el is not None:
                el.text = val
                modificado = True

    return modificado

# ─────────────────────────────────────────────
# PROCESSAMENTO XML
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
    diferencas = []

    for det, n_item, dados in itens_validos:
        prod    = find(det, "prod")
        imposto = find(det, "imposto")
        if prod is None or imposto is None:
            continue

        # ── Coleta ANTES ──
        def _get_txt(elem, *path):
            el = find(elem, *path)
            return (el.text or "0").strip() if el is not None else "0"

        def _fv(v):
            try: return round(float(v), 2)
            except: return 0.0

        antes = {}
        icms_pai = find(imposto, "ICMS")
        if icms_pai:
            for icms_f in icms_pai:
                if local(icms_f).startswith("ICMS"):
                    antes["CST ICMS"]    = _get_txt(icms_f, "CST")
                    antes["BC ICMS"]     = _get_txt(icms_f, "vBC")
                    antes["% ICMS"]      = _get_txt(icms_f, "pICMS")
                    antes["Vlr ICMS"]    = _get_txt(icms_f, "vICMS")
                    antes["BC ICMS ST"]  = _get_txt(icms_f, "vBCST")
                    antes["Vlr ICMS ST"] = _get_txt(icms_f, "vICMSST")
                    break
        ipi_t = find(imposto, "IPI", "IPITrib")
        if ipi_t:
            antes["CST IPI"]  = _get_txt(ipi_t, "CST")
            antes["BC IPI"]   = _get_txt(ipi_t, "vBC")
            antes["% IPI"]    = _get_txt(ipi_t, "pIPI")
            antes["Vlr IPI"]  = _get_txt(ipi_t, "vIPI")
        pis_p = find(imposto, "PIS")
        if pis_p:
            for pf in pis_p:
                antes["CST PIS"]  = _get_txt(pf, "CST")
                antes["BC PIS"]   = _get_txt(pf, "vBC")
                antes["% PIS"]    = _get_txt(pf, "pPIS")
                antes["Vlr PIS"]  = _get_txt(pf, "vPIS")
                break
        cof_p = find(imposto, "COFINS")
        if cof_p:
            for cf in cof_p:
                antes["CST COFINS"]  = _get_txt(cf, "CST")
                antes["BC COFINS"]   = _get_txt(cf, "vBC")
                antes["% COFINS"]    = _get_txt(cf, "pCOFINS")
                antes["Vlr COFINS"]  = _get_txt(cf, "vCOFINS")
                break

        # ── CFOP ──
        el = find(prod, "CFOP")
        cfop_antes = el.text if el is not None else ""
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
                for tn, tv in [
                    ("CST",  dados["cst_ipi"]),
                    ("vBC",  fmt(dados["base_ipi"])),
                    ("pIPI", fmt(dados["perc_ipi"])),
                    ("vIPI", fmt(dados["vlr_ipi"])),
                ]:
                    el = find(ipi_trib, tn)
                    if el is not None:
                        el.text = tv
                        modificado = True

        # ── PIS ──
        pis_pai = find(imposto, "PIS")
        if pis_pai is not None:
            for pf in pis_pai:
                for tn, tv in [
                    ("CST",  dados["cst_pis"]),
                    ("vBC",  fmt(dados["bc_pis_cofins"])),
                    ("pPIS", fmt(dados["perc_pis"])),
                    ("vPIS", fmt(dados["vlr_pis"])),
                ]:
                    el = find(pf, tn)
                    if el is not None:
                        el.text = tv
                        modificado = True
                break

        # ── COFINS ──
        cof_pai = find(imposto, "COFINS")
        if cof_pai is not None:
            for cf in cof_pai:
                for tn, tv in [
                    ("CST",     dados["cst_cofins"]),
                    ("vBC",     fmt(dados["bc_pis_cofins"])),
                    ("pCOFINS", fmt(dados["perc_cofins"])),
                    ("vCOFINS", fmt(dados["vlr_cofins"])),
                ]:
                    el = find(cf, tn)
                    if el is not None:
                        el.text = tv
                        modificado = True
                break

        # ── Coleta DEPOIS ──
        depois = {
            "CST ICMS":    dados["cst_icms"],
            "BC ICMS":     fmt(dados["base_icms"]),
            "% ICMS":      fmt(dados["perc_icms"]),
            "Vlr ICMS":    fmt(dados["vlr_icms"]),
            "BC ICMS ST":  fmt(dados["base_icms_st"]),
            "Vlr ICMS ST": fmt(max(0.0, dados["vlr_icms_st"])),
            "CST IPI":     dados["cst_ipi"],
            "BC IPI":      fmt(dados["base_ipi"]),
            "% IPI":       fmt(dados["perc_ipi"]),
            "Vlr IPI":     fmt(dados["vlr_ipi"]),
            "CST PIS":     dados["cst_pis"],
            "BC PIS":      fmt(dados["bc_pis_cofins"]),
            "% PIS":       fmt(dados["perc_pis"]),
            "Vlr PIS":     fmt(dados["vlr_pis"]),
            "CST COFINS":  dados["cst_cofins"],
            "BC COFINS":   fmt(dados["bc_pis_cofins"]),
            "% COFINS":    fmt(dados["perc_cofins"]),
            "Vlr COFINS":  fmt(dados["vlr_cofins"]),
        }

        # Detecta diferença
        campos_num = [
            "BC ICMS","Vlr ICMS","BC ICMS ST","Vlr ICMS ST",
            "BC IPI","Vlr IPI","BC PIS","Vlr PIS","BC COFINS","Vlr COFINS"
        ]
        tem_diff = any(
            round(_fv(antes.get(k, "0")), 2) != round(_fv(depois.get(k, "0")), 2)
            for k in campos_num
        ) or cfop_antes != dados["cfop"]

        xprod_el = find(prod, "xProd")
        xprod = xprod_el.text if xprod_el is not None else ""

        row_conf = {
            # Identificação
            "Arquivo":        nome_arquivo,
            "Chave NF-e":     chave_xml,
            "Nro Documento":  dados["nro_documento"],
            "Data Emissão":   dados["data_emissao"],
            "Data Entrada":   dados["data_entrada"],
            "Razão Social":   dados["razao_social"],
            "CNPJ/CPF":       dados["cnpj"],
            "UF":             dados["uf"],
            "nItem":          n_item,
            "Cód Item":       dados["cod_item"],
            "Desc Item":      xprod,
            "NCM":            dados["ncm"],
            "CFOP Antes":     cfop_antes,
            "CFOP Depois":    dados["cfop"],
            "Vlr Documento":  fmt(dados["vlr_documento"]),
            # ICMS
            "CST ICMS Antes":    antes.get("CST ICMS", ""),
            "CST ICMS Depois":   depois["CST ICMS"],
            "BC ICMS Antes":     antes.get("BC ICMS", "0"),
            "BC ICMS Depois":    depois["BC ICMS"],
            "% ICMS":            depois["% ICMS"],
            "Vlr ICMS Antes":    antes.get("Vlr ICMS", "0"),
            "Vlr ICMS Depois":   depois["Vlr ICMS"],
            "Diff Vlr ICMS":     round(_fv(depois["Vlr ICMS"]) - _fv(antes.get("Vlr ICMS","0")), 2),
            # ICMS ST
            "BC ICMS ST Antes":  antes.get("BC ICMS ST", "0"),
            "BC ICMS ST Depois": depois["BC ICMS ST"],
            "Vlr ICMS ST Antes": antes.get("Vlr ICMS ST", "0"),
            "Vlr ICMS ST Depois":depois["Vlr ICMS ST"],
            "Diff Vlr ICMS ST":  round(_fv(depois["Vlr ICMS ST"]) - _fv(antes.get("Vlr ICMS ST","0")), 2),
            # IPI
            "CST IPI Antes":     antes.get("CST IPI", ""),
            "CST IPI Depois":    depois["CST IPI"],
            "BC IPI Antes":      antes.get("BC IPI", "0"),
            "BC IPI Depois":     depois["BC IPI"],
            "% IPI":             depois["% IPI"],
            "Vlr IPI Antes":     antes.get("Vlr IPI", "0"),
            "Vlr IPI Depois":    depois["Vlr IPI"],
            "Diff Vlr IPI":      round(_fv(depois["Vlr IPI"]) - _fv(antes.get("Vlr IPI","0")), 2),
            # PIS
            "CST PIS Antes":     antes.get("CST PIS", ""),
            "CST PIS Depois":    depois["CST PIS"],
            "BC PIS Antes":      antes.get("BC PIS", "0"),
            "BC PIS Depois":     depois["BC PIS"],
            "% PIS":             depois["% PIS"],
            "Vlr PIS Antes":     antes.get("Vlr PIS", "0"),
            "Vlr PIS Depois":    depois["Vlr PIS"],
            "Diff Vlr PIS":      round(_fv(depois["Vlr PIS"]) - _fv(antes.get("Vlr PIS","0")), 2),
            # COFINS
            "CST COFINS Antes":  antes.get("CST COFINS", ""),
            "CST COFINS Depois": depois["CST COFINS"],
            "BC COFINS Antes":   antes.get("BC COFINS", "0"),
            "BC COFINS Depois":  depois["BC COFINS"],
            "% COFINS":          depois["% COFINS"],
            "Vlr COFINS Antes":  antes.get("Vlr COFINS", "0"),
            "Vlr COFINS Depois": depois["Vlr COFINS"],
            "Diff Vlr COFINS":   round(_fv(depois["Vlr COFINS"]) - _fv(antes.get("Vlr COFINS","0")), 2),
            # Flag
            "Tem Diferença":     "SIM" if tem_diff else "NÃO",
        }
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

    for tag_nome, val in {
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
    }.items():
        el = find(icms_tot, tag_nome)
        if el is not None:
            el.text = val


# ─────────────────────────────────────────────
# EXCEL DE CONFERÊNCIA
# ─────────────────────────────────────────────
def gerar_excel_conferencia(todas_diferencas: list) -> bytes:
    if not todas_diferencas:
        return b""

    df = pd.DataFrame(todas_diferencas)
    buf = io.BytesIO()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Conferência")
        ws = writer.sheets["Conferência"]

        from openpyxl.styles import (
            PatternFill, Font, Alignment, Border, Side, GradientFill
        )
        from openpyxl.utils import get_column_letter

        # Cores tema TR
        COR_LARANJA   = "FF8000"
        COR_CINZA     = "444444"
        COR_BRANCO    = "FFFFFF"
        COR_VERDE_BG  = "E2EFDA"
        COR_VERDE_FT  = "375623"
        COR_VERM_BG   = "FCE4D6"
        COR_VERM_FT   = "843C0C"
        COR_AMARELO   = "FFF2CC"
        COR_LINHA_PAR = "F5F5F5"

        thin  = Side(style="thin",   color="CCCCCC")
        thick = Side(style="medium", color=COR_LARANJA)
        borda_header = Border(left=thick, right=thick, top=thick, bottom=thick)
        borda_normal = Border(left=thin,  right=thin,  top=thin,  bottom=thin)

        header_fill = PatternFill("solid", fgColor=COR_CINZA)
        header_font = Font(color=COR_BRANCO, bold=True, size=9,
                           name="Segoe UI")
        header_align = Alignment(horizontal="center", vertical="center",
                                 wrap_text=True)

        # Grupos de colunas para colorir o cabeçalho
        grupos = {
            "identificacao": {
                "cols": ["Arquivo","Chave NF-e","Nro Documento","Data Emissão",
                         "Data Entrada","Razão Social","CNPJ/CPF","UF",
                         "nItem","Cód Item","Desc Item","NCM",
                         "CFOP Antes","CFOP Depois","Vlr Documento"],
                "cor": "2E4057",
            },
            "icms": {
                "cols": ["CST ICMS Antes","CST ICMS Depois","BC ICMS Antes",
                         "BC ICMS Depois","% ICMS","Vlr ICMS Antes",
                         "Vlr ICMS Depois","Diff Vlr ICMS",
                         "BC ICMS ST Antes","BC ICMS ST Depois",
                         "Vlr ICMS ST Antes","Vlr ICMS ST Depois",
                         "Diff Vlr ICMS ST"],
                "cor": "1F4E79",
            },
            "ipi": {
                "cols": ["CST IPI Antes","CST IPI Depois","BC IPI Antes",
                         "BC IPI Depois","% IPI","Vlr IPI Antes",
                         "Vlr IPI Depois","Diff Vlr IPI"],
                "cor": "375623",
            },
            "pis": {
                "cols": ["CST PIS Antes","CST PIS Depois","BC PIS Antes",
                         "BC PIS Depois","% PIS","Vlr PIS Antes",
                         "Vlr PIS Depois","Diff Vlr PIS"],
                "cor": "7B2C2C",
            },
            "cofins": {
                "cols": ["CST COFINS Antes","CST COFINS Depois","BC COFINS Antes",
                         "BC COFINS Depois","% COFINS","Vlr COFINS Antes",
                         "Vlr COFINS Depois","Diff Vlr COFINS"],
                "cor": "843C0C",
            },
            "flag": {
                "cols": ["Tem Diferença"],
                "cor": COR_LARANJA,
            },
        }

        col_names = list(df.columns)
        col_grupo = {}
        for grp, info in grupos.items():
            for c in info["cols"]:
                if c in col_names:
                    col_grupo[col_names.index(c) + 1] = info["cor"]

        # Aplica cabeçalho
        for cell in ws[1]:
            cor = col_grupo.get(cell.column, COR_CINZA)
            cell.fill      = PatternFill("solid", fgColor=cor)
            cell.font      = Font(color=COR_BRANCO, bold=True, size=9, name="Segoe UI")
            cell.alignment = header_align
            cell.border    = borda_header

        ws.row_dimensions[1].height = 36

        # Colunas diff e flag
        diff_cols = [i+1 for i, c in enumerate(col_names) if c.startswith("Diff ")]
        flag_col  = col_names.index("Tem Diferença") + 1 if "Tem Diferença" in col_names else None

        fill_par = PatternFill("solid", fgColor=COR_LINHA_PAR)

        for row_idx, row in enumerate(
            ws.iter_rows(min_row=2, max_row=ws.max_row), start=2
        ):
            is_par = (row_idx % 2 == 0)
            for cell in row:
                cell.border    = borda_normal
                cell.alignment = Alignment(horizontal="center", vertical="center",
                                           wrap_text=False)
                cell.font      = Font(size=9, name="Segoe UI")
                if is_par:
                    cell.fill = fill_par

            # Destaca diffs
            for col_idx in diff_cols:
                cell = row[col_idx - 1]
                try:
                    v = float(cell.value or 0)
                    if v > 0:
                        cell.fill = PatternFill("solid", fgColor=COR_VERDE_BG)
                        cell.font = Font(color=COR_VERDE_FT, bold=True, size=9, name="Segoe UI")
                    elif v < 0:
                        cell.fill = PatternFill("solid", fgColor=COR_VERM_BG)
                        cell.font = Font(color=COR_VERM_FT, bold=True, size=9, name="Segoe UI")
                except Exception:
                    pass

            # Flag "Tem Diferença"
            if flag_col:
                cell = row[flag_col - 1]
                if str(cell.value) == "SIM":
                    cell.fill = PatternFill("solid", fgColor="FFE0B2")
                    cell.font = Font(color=COR_LARANJA, bold=True, size=9, name="Segoe UI")
                else:
                    cell.fill = PatternFill("solid", fgColor=COR_VERDE_BG)
                    cell.font = Font(color=COR_VERDE_FT, bold=True, size=9, name="Segoe UI")

        # Largura das colunas
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            header_val = str(ws.cell(1, col_idx).value or "")
            max_len = len(header_val)
            for row_idx in range(2, min(ws.max_row + 1, 100)):
                v = str(ws.cell(row_idx, col_idx).value or "")
                max_len = max(max_len, len(v))
            ws.column_dimensions[col_letter].width = min(max_len + 3, 35)

        # Colunas de descrição mais largas
        for col_name in ["Desc Item", "Razão Social", "Arquivo", "Chave NF-e"]:
            if col_name in col_names:
                ws.column_dimensions[
                    get_column_letter(col_names.index(col_name) + 1)
                ].width = 40

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# INTERFACE
# ─────────────────────────────────────────────

st.markdown('<div class="section-card"><div class="section-title">📂 1. Arquivo de Entrada (XLSX — Domínio Sistemas)</div>', unsafe_allow_html=True)
arquivo_xlsx = st.file_uploader("Selecione o arquivo .xlsx exportado do Domínio", type=["xlsx"], key="xlsx")
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-card"><div class="section-title">📄 2. Arquivos XML de NF-e</div>', unsafe_allow_html=True)
arquivos_xml = st.file_uploader(
    "Selecione XMLs ou um arquivo ZIP contendo XMLs",
    type=["xml", "zip"],
    accept_multiple_files=True,
    key="xmls"
)
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-card"><div class="section-title">🔢 3. CFOPs a Processar</div>', unsafe_allow_html=True)

CFOPS_DEFAULT = ["1201","1202","1410","1411","2201","2202","2410","2411","1949","2949","2603"]

if "cfops_lista" not in st.session_state:
    st.session_state["cfops_lista"] = list(CFOPS_DEFAULT)

col_inp, col_add = st.columns([3, 1])
with col_inp:
    nova_cfop = st.text_input(
        "CFOP", max_chars=4, placeholder="ex: 1410",
        key="nova_cfop_input", label_visibility="collapsed"
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

if st.session_state["cfops_lista"]:
    cfops_para_remover = []
    n_cols = min(len(st.session_state["cfops_lista"]), 10)
    cols_cfop = st.columns(n_cols)
    for i, cfop in enumerate(st.session_state["cfops_lista"]):
        with cols_cfop[i % n_cols]:
            if st.button(f"❌ {cfop}", key=f"rm_{cfop}_{i}", help=f"Remover {cfop}"):
                cfops_para_remover.append(cfop)
    if cfops_para_remover:
        for c in cfops_para_remover:
            if c in st.session_state["cfops_lista"]:
                st.session_state["cfops_lista"].remove(c)
        st.rerun()

col_lote, col_reset = st.columns([3, 1])
with col_lote:
    remover_lote = st.text_input(
        "Remover em lote", placeholder="ex: 1949, 2949",
        key="remover_lote_input", label_visibility="collapsed"
    )
with col_reset:
    if st.button("🔄 Restaurar padrão", key="btn_reset"):
        st.session_state["cfops_lista"] = list(CFOPS_DEFAULT)
        st.rerun()

if st.button("🗑️ Remover em lote", key="btn_rm_lote"):
    for c in [re.sub(r"[^0-9]", "", x.strip()) for x in remover_lote.split(",") if x.strip()]:
        if c in st.session_state["cfops_lista"]:
            st.session_state["cfops_lista"].remove(c)
    st.rerun()

st.caption(f"CFOPs ativas: **{', '.join(sorted(st.session_state['cfops_lista'])) or 'Nenhuma'}**")
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-card"><div class="section-title">▶️ 4. Processar</div>', unsafe_allow_html=True)

if st.button("▶ Processar XMLs", type="primary", key="btn_processar"):
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
        st.error("❌ Nenhum item válido no XLSX.")
        st.stop()

    col_a, col_b, col_c = st.columns(3)
    col_a.metric("📊 Itens indexados", len(dados_indexados))
    col_b.metric("🔢 CFOPs ativas",    len(cfops_ativas))
    col_c.metric("📁 Arquivos",        len(arquivos_xml))

    with st.expander("🔍 Debug — primeiros 5 itens indexados"):
        for i, ((ch, seq), d) in enumerate(list(dados_indexados.items())[:5]):
            st.code(
                f"Chave: {ch} | Seq: {seq} | CFOP: {d['cfop']}\n"
                f"Item: {d['cod_item']} | vICMS: {d['vlr_icms']} | "
                f"vICMSST: {d['vlr_icms_st']} | vIPI: {d['vlr_ipi']} | "
                f"BC PIS/COFINS: {d['bc_pis_cofins']}"
            )

    # Coleta XMLs
    xmls_para_processar = {}
    for arq in arquivos_xml:
        if arq.name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(arq.read())) as zf:
                for nome in zf.namelist():
                    if nome.lower().endswith(".xml"):
                        xmls_para_processar[nome] = zf.read(nome)
        else:
            xmls_para_processar[arq.name] = arq.read()

    resultados       = []
    xmls_modificados = {}
    todas_diferencas = []
    progress = st.progress(0)
    total    = len(xmls_para_processar)

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

    ok_c   = sum(1 for _, _, s in resultados if s == "ok")
    info_c = sum(1 for _, _, s in resultados if s == "info")
    err_c  = sum(1 for _, _, s in resultados if s == "erro")
    diff_c = sum(1 for r in todas_diferencas if r.get("Tem Diferença") == "SIM")

    col_x, col_y, col_z, col_w = st.columns(4)
    col_x.metric("✅ Alterados",      ok_c)
    col_y.metric("ℹ️ Sem alteração",  info_c)
    col_z.metric("❌ Erros",          err_c)
    col_w.metric("📋 Itens c/ diff",  diff_c)

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
        st.subheader("⬇️ Download XMLs Alterados")
        if len(xmls_modificados) == 1:
            nome_arq, conteudo_arq = list(xmls_modificados.items())[0]
            st.download_button(
                label=f"⬇ Baixar {nome_arq}",
                data=conteudo_arq,
                file_name=nome_arq,
                mime="application/xml",
                use_container_width=True,
            )
        else:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for nome_arq, conteudo_arq in xmls_modificados.items():
                    zf.writestr(nome_arq, conteudo_arq)
            buf.seek(0)
            st.download_button(
                label=f"⬇ Baixar todos ({len(xmls_modificados)} XMLs) como ZIP",
                data=buf,
                file_name="xmls_modificados_dni.zip",
                mime="application/zip",
                use_container_width=True,
            )

    # Download Excel de Conferência
    if todas_diferencas:
        st.subheader("📊 Excel de Conferência")
        excel_bytes = gerar_excel_conferencia(todas_diferencas)
        st.download_button(
            label=f"⬇ Baixar Excel de Conferência ({len(todas_diferencas)} itens processados / {diff_c} com diferença)",
            data=excel_bytes,
            file_name="conferencia_nfe_dni.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.info("ℹ️ Nenhum item processado para conferência.")

    if not xmls_modificados:
        st.warning("⚠️ Nenhum XML foi modificado.")

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("""
<div class="footer">
    Thomson Reuters · Domínio Sistemas · Enriquecedor NF-e v4.0 · DNI
</div>
""", unsafe_allow_html=True)
