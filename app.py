import streamlit as st
import pandas as pd
import io
import re
import zipfile
from lxml import etree

# ─────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Enriquecedor de NF-e — DNI",
    page_icon="🟠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CATÁLOGO PADRÃO DE CFOPs (imutável)
# ─────────────────────────────────────────────
CATALOGO_PADRAO: dict[int, str] = {
    1201: "Dev. de Venda de Prod. do Estabelecimento — Estadual",
    1202: "Dev. de Venda de Merc. Adq./Rec. de Terceiros — Estadual",
    1410: "Devolução de Venda Própria ST — Estadual",
    1411: "Devolução de Venda Terceiros ST — Estadual",
    1949: "Outra Entrada de Mercadoria — Estadual",
    2201: "Dev. de Venda de Prod. do Estabelecimento — Interestadual",
    2202: "Dev. de Venda de Merc. Adq./Rec. de Terceiros — Interestadual",
    2410: "Devolução de Venda Própria ST — Interestadual",
    2411: "Devolução de Venda Terceiros ST — Interestadual",
    2603: "Devolução de Venda de Energia Elétrica — Interestadual",
    2949: "Outra Entrada de Mercadoria — Interestadual",
}
ATIVOS_PADRAO: set[int] = set(CATALOGO_PADRAO.keys())

# ─────────────────────────────────────────────
# TEMA THOMSON REUTERS
# ─────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: 'Segoe UI', 'Arial', sans-serif;
    color: #444444;
}
h1, h2, h3 { color: #FF8000; font-weight: 700; }
.main-header {
    background: #444444;
    padding: 18px 28px 14px 28px;
    border-radius: 8px;
    border-top: 6px solid #FF8000;
    margin-bottom: 20px;
}
.main-header h2 { color: #FF8000 !important; margin: 0; font-size: 1.4rem; }
.main-header p  { color: #DDDDDD !important; margin: 4px 0 0 0; font-size: 0.82rem; }
section[data-testid="stSidebar"] { background-color: #2E2E2E; }
section[data-testid="stSidebar"] * { color: #FFFFFF !important; }
section[data-testid="stSidebar"] .stButton > button {
    background-color: #FF8000 !important; color: #FFFFFF !important;
    border: none !important; border-radius: 4px !important; font-weight: bold !important;
}
section[data-testid="stSidebar"] .stButton > button:hover { background-color: #D64001 !important; }
.stButton > button {
    background-color: #FF8000 !important; color: #FFFFFF !important;
    border: none !important; border-radius: 4px !important; font-weight: bold !important;
}
.stButton > button:hover { background-color: #D64001 !important; }
.stDownloadButton > button {
    background-color: #FF8000 !important; color: #FFFFFF !important;
    border: none !important; border-radius: 4px !important; font-weight: bold !important;
}
.stDownloadButton > button:hover { background-color: #D64001 !important; }
[data-testid="metric-container"] {
    background-color: #F5F5F5; border-left: 4px solid #FF8000;
    border-radius: 4px; padding: 8px 12px;
}
.footer {
    text-align: center; color: #999; font-size: 0.72rem;
    margin-top: 24px; padding-top: 12px; border-top: 1px solid #E0E0E0;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
def _init_state():
    if "cfop_catalogo" not in st.session_state:
        st.session_state.cfop_catalogo = dict(CATALOGO_PADRAO)
    if "cfop_ativos" not in st.session_state:
        st.session_state.cfop_ativos = set(ATIVOS_PADRAO)
    if "resultado_excel_bytes" not in st.session_state:
        st.session_state.resultado_excel_bytes = None
    if "resultado_xmls" not in st.session_state:
        st.session_state.resultado_xmls = {}
    if "resultado_log" not in st.session_state:
        st.session_state.resultado_log = []
    if "resultado_metricas" not in st.session_state:
        st.session_state.resultado_metricas = None
    if "resultado_diferencas" not in st.session_state:
        st.session_state.resultado_diferencas = []
    if "processamento_concluido" not in st.session_state:
        st.session_state.processamento_concluido = False

def _limpar_resultados():
    st.session_state.resultado_excel_bytes = None
    st.session_state.resultado_xmls = {}
    st.session_state.resultado_log = []
    st.session_state.resultado_metricas = None
    st.session_state.resultado_diferencas = []
    st.session_state.processamento_concluido = False

# ─────────────────────────────────────────────
# HELPERS CFOP
# ─────────────────────────────────────────────
def _extrair_cfops_texto(texto: str) -> list[int]:
    tokens = re.findall(r"\d+", texto)
    return [int(t) for t in tokens if len(t) == 4]

def _build_rows() -> list[dict]:
    catalogo = st.session_state.cfop_catalogo
    ativos   = st.session_state.cfop_ativos
    return [
        {"Ativo": cod in ativos, "Código": cod, "Descrição": desc}
        for cod, desc in sorted(catalogo.items())
    ]

# ─────────────────────────────────────────────
# BARRA LATERAL
# ─────────────────────────────────────────────
def render_sidebar():
    with st.sidebar:
        st.markdown("## ⚙️ Configuração de CFOPs")
        st.markdown("---")
        st.markdown("### 📥 Entrada em Lote")
        st.markdown(
            "<small>Cole CFOPs separados por espaço, vírgula ou quebra de linha.</small>",
            unsafe_allow_html=True,
        )
        texto = st.text_area(
            label="CFOPs em lote",
            placeholder="Ex:\n1201, 1410\n2411 2603",
            height=130,
            key="sidebar_cfop_texto",
            label_visibility="collapsed",
        )
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("➕ Ativar", use_container_width=True, key="btn_ativar_lote"):
                codigos = _extrair_cfops_texto(texto)
                if codigos:
                    for c in codigos:
                        if c not in st.session_state.cfop_catalogo:
                            st.session_state.cfop_catalogo[c] = "CFOP Personalizado"
                        st.session_state.cfop_ativos.add(c)
                    st.toast(f"✅ {len(codigos)} CFOP(s) ativado(s).", icon="✅")
                    st.rerun()
                else:
                    st.warning("Nenhum CFOP válido encontrado.")
        with col_b:
            if st.button("🗑️ Desativar", use_container_width=True, key="btn_desativar_lote"):
                codigos = _extrair_cfops_texto(texto)
                if codigos:
                    for c in codigos:
                        st.session_state.cfop_ativos.discard(c)
                    st.toast("🗑️ CFOPs desativados.", icon="🗑️")
                    st.rerun()
                else:
                    st.warning("Nenhum CFOP válido encontrado.")

        st.markdown("---")
        st.markdown("### 🎛️ Controles")
        if st.button("🔄 Redefinir para o Padrão", use_container_width=True, key="btn_reset_cfop"):
            st.session_state.cfop_catalogo = dict(CATALOGO_PADRAO)
            st.session_state.cfop_ativos   = set(ATIVOS_PADRAO)
            st.toast("✅ Catálogo restaurado para o padrão.", icon="🔄")
            st.rerun()
        if st.button("🔲 Desmarcar Todos", use_container_width=True, key="btn_desmarcar_todos"):
            st.session_state.cfop_ativos = set()
            st.toast("🔲 Todos os CFOPs desmarcados.", icon="🔲")
            st.rerun()

        st.markdown("---")
        st.markdown("### 🗂️ CFOPs Cadastrados")
        rows   = _build_rows()
        edited = st.data_editor(
            rows,
            column_config={
                "Ativo": st.column_config.CheckboxColumn(
                    label="Ativo", help="Marque para incluir no processamento.", default=False,
                ),
                "Código": st.column_config.NumberColumn(label="Código", format="%d", disabled=True),
                "Descrição": st.column_config.TextColumn(label="Descrição", disabled=True),
            },
            hide_index=True,
            use_container_width=True,
            key="cfop_data_editor",
        )
        novos_ativos: set[int] = {row["Código"] for row in edited if row["Ativo"]}
        if novos_ativos != st.session_state.cfop_ativos:
            st.session_state.cfop_ativos = novos_ativos
            st.rerun()

        st.markdown("---")
        st.markdown("#### 🐍 Array Python (backend)")
        lista_ativos = sorted(st.session_state.cfop_ativos)
        st.code(str(lista_ativos), language="python")
        total  = len(st.session_state.cfop_catalogo)
        ativos = len(lista_ativos)
        st.caption(f"Catálogo: **{total}** | Ativos: **{ativos}** | Inativos: **{total - ativos}**")

        st.markdown("---")
        st.markdown("### ℹ️ Sobre")
        st.markdown("**Thomson Reuters**")
        st.markdown("**Domínio Sistemas**")
        st.markdown("**Enriquecedor NF-e v7.3**")

# ─────────────────────────────────────────────
# HELPERS GERAIS
# ─────────────────────────────────────────────
def limpar_chave(valor: str) -> str:
    if not valor:
        return ""
    return re.sub(r"[^0-9]", "", str(valor).strip())

def limpar_valor(valor) -> float:
    if valor is None:
        return 0.0
    try:
        if pd.isna(valor):
            return 0.0
    except Exception:
        pass
    s = re.sub(r"[^\d\.\-]", "", str(valor).strip().replace(",", "."))
    try:
        return float(s)
    except ValueError:
        return 0.0

def fmt(v: float) -> str:
    return f"{v:.2f}"

def fmt_br(v) -> str:
    try:
        return f"{float(v):.2f}".replace(".", ",")
    except Exception:
        return str(v)

def safe_int_cst(val) -> str:
    if val is None:
        return "00"
    try:
        if pd.isna(val):
            return "00"
    except Exception:
        pass
    s = str(val).strip()
    if s.lower() in ("nan", "none", ""):
        return "00"
    try:
        return str(int(float(s))).zfill(2)
    except Exception:
        return "00"

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

        cst_icms = safe_int_cst(reg.get("CST ICMS", ""))
        if cst_icms == "00" and (
            limpar_valor(reg.get("Base Icms St", "0")) > 0 or
            limpar_valor(reg.get("Vlr Icms St",  "0")) > 0
        ):
            cst_icms = "10"

        base_ipi = limpar_valor(reg.get("Base Ipi", "0"))
        perc_ipi = limpar_valor(reg.get("Perc Ipi", "0"))
        tem_ipi  = (base_ipi > 0 or vlr_ipi > 0)

        indexado[(chave_nfe, seq)] = {
            # ── CFOP mantido apenas para conferência — NÃO aplicado no XML ──
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
            "cst_icms":       cst_icms,
            "base_icms":      limpar_valor(reg.get("Base Icms", "0")),
            "perc_icms":      limpar_valor(reg.get("Perc ICms", "0")),
            "vlr_icms":       vlr_icms,
            "base_icms_st":   limpar_valor(reg.get("Base Icms St", "0")),
            "vlr_icms_st":    vlr_icms_st,
            "cst_ipi":        safe_int_cst(reg.get("CST IPI", "")),
            "base_ipi":       base_ipi,
            "perc_ipi":       perc_ipi,
            "vlr_ipi":        vlr_ipi,
            "tem_ipi":        tem_ipi,
            "cst_pis":        safe_int_cst(reg.get("CST PIS", "")),
            "perc_pis":       limpar_valor(reg.get("Perc Pis", "0")),
            "vlr_pis":        limpar_valor(reg.get("Vlr Pis", "0")),
            "cst_cofins":     safe_int_cst(reg.get("CST COFINS", "")),
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

def nstag(nome):  return f"{{{NS}}}{nome}"
def local(elem):  return elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

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

def _insert_after(parent, ref_tag_name: str, new_elem):
    for i, child in enumerate(parent):
        if local(child) == ref_tag_name:
            parent.insert(i + 1, new_elem)
            return True
    parent.append(new_elem)
    return False

# ─────────────────────────────────────────────
# ICMS
# ─────────────────────────────────────────────
CST_COM_ST = {"10", "30", "70", "90"}

def aplicar_icms(icms_filho, dados):
    cst = dados["cst_icms"]
    modificado = False

    el = find(icms_filho, "CST")
    if el is not None:
        el.text = cst
        modificado = True

    for tn, tv in [
        ("vBC",   fmt(dados["base_icms"])),
        ("pICMS", fmt(dados["perc_icms"])),
        ("vICMS", fmt(dados["vlr_icms"])),
    ]:
        el = find(icms_filho, tn)
        if el is not None:
            el.text = tv
            modificado = True

    tem_st = (
        cst in CST_COM_ST and
        (dados["base_icms_st"] > 0 or dados["vlr_icms_st"] > 0)
    )

    if tem_st:
        el_mod = find(icms_filho, "modBCST")
        if el_mod is None:
            idx_vicms = next(
                (i for i, f in enumerate(icms_filho) if local(f) == "vICMS"), None
            )
            el_mod = etree.SubElement(icms_filho, nstag("modBCST"))
            el_mod.text = "4"
            if idx_vicms is not None:
                icms_filho.remove(el_mod)
                icms_filho.insert(idx_vicms + 1, el_mod)

        mod_val = (el_mod.text or "4").strip()
        if mod_val == "4":
            el_pmva = find(icms_filho, "pMVAST")
            if el_pmva is None:
                idx_mod = next(
                    (i for i, f in enumerate(icms_filho) if local(f) == "modBCST"), None
                )
                el_pmva = etree.SubElement(icms_filho, nstag("pMVAST"))
                el_pmva.text = "0.00"
                if idx_mod is not None:
                    icms_filho.remove(el_pmva)
                    icms_filho.insert(idx_mod + 1, el_pmva)

        for tn, tv in [
            ("vBCST",   fmt(dados["base_icms_st"])),
            ("pICMSST", fmt(dados["perc_icms"])),
            ("vICMSST", fmt(max(0.0, dados["vlr_icms_st"]))),
        ]:
            el = find(icms_filho, tn)
            if el is not None:
                el.text = tv
                modificado = True
            else:
                novo = etree.SubElement(icms_filho, nstag(tn))
                novo.text = tv
                modificado = True

    return modificado

# ─────────────────────────────────────────────
# IPI — gera vBC/pIPI/vIPI e zera impostoDevol
# ─────────────────────────────────────────────
def aplicar_ipi(imposto_elem, det_elem, dados):
    """
    1. Se XLSX tem IPI (tem_ipi=True):
       - Se <IPI> não existe em <imposto>: cria bloco completo antes de <PIS>
       - Se <IPI> existe: remove qUnid/vUnid, cria/atualiza vBC, pIPI, vIPI
    2. Zera <pDevol> e <vIPIDevol> em <impostoDevol> — SEMPRE
    """
    modificado = False
    tem_ipi    = dados.get("tem_ipi", False)

    if tem_ipi:
        ipi_elem = find(imposto_elem, "IPI")

        # ── Bloco IPI não existe → cria completo ──
        if ipi_elem is None:
            ipi_elem = etree.Element(nstag("IPI"))

            el_cenq      = etree.SubElement(ipi_elem, nstag("cEnq"))
            el_cenq.text = "999"

            ipi_trib     = etree.SubElement(ipi_elem, nstag("IPITrib"))

            el_cst       = etree.SubElement(ipi_trib, nstag("CST"))
            el_cst.text  = dados["cst_ipi"] if dados["cst_ipi"] != "00" else "50"

            el_vbc       = etree.SubElement(ipi_trib, nstag("vBC"))
            el_vbc.text  = fmt(dados["base_ipi"])

            el_pipi      = etree.SubElement(ipi_trib, nstag("pIPI"))
            el_pipi.text = fmt(dados["perc_ipi"])

            el_vipi      = etree.SubElement(ipi_trib, nstag("vIPI"))
            el_vipi.text = fmt(dados["vlr_ipi"])

            pis_elem = find(imposto_elem, "PIS")
            if pis_elem is not None:
                idx_pis = list(imposto_elem).index(pis_elem)
                imposto_elem.insert(idx_pis, ipi_elem)
            else:
                imposto_elem.append(ipi_elem)

            modificado = True

        # ── Bloco IPI já existe → atualiza ──
        else:
            ipi_trib = find(ipi_elem, "IPITrib")
            if ipi_trib is not None:

                # CST
                el_cst = find(ipi_trib, "CST")
                if el_cst is not None:
                    el_cst.text = dados["cst_ipi"]
                    modificado  = True

                # Remove qUnid / vUnid
                for tag_rem in ["qUnid", "vUnid"]:
                    el_rem = find(ipi_trib, tag_rem)
                    if el_rem is not None:
                        ipi_trib.remove(el_rem)
                        modificado = True

                # vBC
                el_vbc = find(ipi_trib, "vBC")
                if el_vbc is None:
                    el_vbc      = etree.Element(nstag("vBC"))
                    el_vbc.text = fmt(dados["base_ipi"])
                    _insert_after(ipi_trib, "CST", el_vbc)
                    modificado  = True
                else:
                    el_vbc.text = fmt(dados["base_ipi"])
                    modificado  = True

                # pIPI
                el_pipi = find(ipi_trib, "pIPI")
                if el_pipi is None:
                    el_pipi      = etree.Element(nstag("pIPI"))
                    el_pipi.text = fmt(dados["perc_ipi"])
                    _insert_after(ipi_trib, "vBC", el_pipi)
                    modificado   = True
                else:
                    el_pipi.text = fmt(dados["perc_ipi"])
                    modificado   = True

                # vIPI
                el_vipi = find(ipi_trib, "vIPI")
                if el_vipi is None:
                    el_vipi      = etree.Element(nstag("vIPI"))
                    el_vipi.text = fmt(dados["vlr_ipi"])
                    _insert_after(ipi_trib, "pIPI", el_vipi)
                    modificado   = True
                else:
                    el_vipi.text = fmt(dados["vlr_ipi"])
                    modificado   = True

    # ── Zera <pDevol> e <vIPIDevol> em <impostoDevol> — SEMPRE ──
    imp_devol = find(det_elem, "impostoDevol")
    if imp_devol is not None:
        # Zera pDevol
        el_pdevol = find(imp_devol, "pDevol")
        if el_pdevol is not None:
            el_pdevol.text = "0.00"
            modificado = True

        # Zera vIPIDevol
        ipi_devol = find(imp_devol, "IPI")
        if ipi_devol is not None:
            el_vipidevol = find(ipi_devol, "vIPIDevol")
            if el_vipidevol is not None:
                el_vipidevol.text = "0.00"
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
        # Filtra pelo CFOP que está no XML — não pelo XLSX
        cfop_xml = ""
        prod_el  = find(det, "prod")
        if prod_el is not None:
            el = find(prod_el, "CFOP")
            if el is not None:
                cfop_xml = (el.text or "").strip()
        if cfop_xml in cfops_ativas:
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

        def _get(elem, *path):
            el = find(elem, *path)
            return (el.text or "0").strip() if el is not None else "0"

        def _fv(v):
            try:
                return round(float(str(v).replace(",", ".")), 2)
            except Exception:
                return 0.0

        # ── Coleta ANTES ──
        antes = {}
        icms_pai = find(imposto, "ICMS")
        if icms_pai:
            for icms_f in icms_pai:
                if local(icms_f).startswith("ICMS"):
                    antes["CST ICMS"]    = _get(icms_f, "CST")
                    antes["BC ICMS"]     = _get(icms_f, "vBC")
                    antes["% ICMS"]      = _get(icms_f, "pICMS")
                    antes["Vlr ICMS"]    = _get(icms_f, "vICMS")
                    antes["BC ICMS ST"]  = _get(icms_f, "vBCST")
                    antes["Vlr ICMS ST"] = _get(icms_f, "vICMSST")
                    break

        ipi_t = find(imposto, "IPI", "IPITrib")
        if ipi_t:
            antes["CST IPI"] = _get(ipi_t, "CST")
            vbc_antes        = _get(ipi_t, "vBC")
            if vbc_antes == "0":
                vbc_antes    = _get(ipi_t, "qUnid")
            antes["BC IPI"]  = vbc_antes
            antes["% IPI"]   = _get(ipi_t, "pIPI")
            antes["Vlr IPI"] = _get(ipi_t, "vIPI")

        pis_p = find(imposto, "PIS")
        if pis_p:
            for pf in pis_p:
                antes["CST PIS"] = _get(pf, "CST")
                antes["BC PIS"]  = _get(pf, "vBC")
                antes["% PIS"]   = _get(pf, "pPIS")
                antes["Vlr PIS"] = _get(pf, "vPIS")
                break

        cof_p = find(imposto, "COFINS")
        if cof_p:
            for cf in cof_p:
                antes["CST COFINS"] = _get(cf, "CST")
                antes["BC COFINS"]  = _get(cf, "vBC")
                antes["% COFINS"]   = _get(cf, "pCOFINS")
                antes["Vlr COFINS"] = _get(cf, "vCOFINS")
                break

        # Coleta impostoDevol ANTES
        imp_devol_el = find(det, "impostoDevol")
        antes["pDevol"]     = _get(imp_devol_el, "pDevol")     if imp_devol_el is not None else "0"
        antes["vIPIDevol"]  = _get(imp_devol_el, "IPI", "vIPIDevol") if imp_devol_el is not None else "0"

        # CFOP do XML (mantido — não alterado)
        cfop_xml_atual = _get(prod, "CFOP")

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

        # ── IPI + impostoDevol ──
        if aplicar_ipi(imposto, det, dados):
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
            "pDevol":      "0.00",
            "vIPIDevol":   "0.00",
        }

        campos_num = [
            "BC ICMS","Vlr ICMS","BC ICMS ST","Vlr ICMS ST",
            "BC IPI","Vlr IPI","BC PIS","Vlr PIS","BC COFINS","Vlr COFINS",
            "vIPIDevol",
        ]
        tem_diff = any(
            round(_fv(antes.get(k, "0")), 2) != round(_fv(depois.get(k, "0")), 2)
            for k in campos_num
        )

        xprod_el = find(prod, "xProd")
        xprod    = xprod_el.text if xprod_el is not None else ""

        row_conf = {
            "Arquivo":              nome_arquivo,
            "Chave NF-e":           chave_xml,
            "Nro Documento":        dados["nro_documento"],
            "Data Emissão":         dados["data_emissao"],
            "Data Entrada":         dados["data_entrada"],
            "Razão Social":         dados["razao_social"],
            "CNPJ/CPF":             dados["cnpj"],
            "UF":                   dados["uf"],
            "nItem":                n_item,
            "Cód Item":             dados["cod_item"],
            "Desc Item":            xprod,
            "NCM":                  dados["ncm"],
            "CFOP XML":             cfop_xml_atual,
            "CFOP XLSX":            dados["cfop"],
            "Vlr Documento":        fmt_br(dados["vlr_documento"]),
            # ICMS
            "CST ICMS Antes":       antes.get("CST ICMS", ""),
            "CST ICMS Depois":      depois["CST ICMS"],
            "BC ICMS Antes":        fmt_br(antes.get("BC ICMS", "0")),
            "BC ICMS Depois":       fmt_br(depois["BC ICMS"]),
            "% ICMS":               fmt_br(depois["% ICMS"]),
            "Vlr ICMS Antes":       fmt_br(antes.get("Vlr ICMS", "0")),
            "Vlr ICMS Depois":      fmt_br(depois["Vlr ICMS"]),
            "Diff Vlr ICMS":        fmt_br(_fv(depois["Vlr ICMS"]) - _fv(antes.get("Vlr ICMS","0"))),
            "BC ICMS ST Antes":     fmt_br(antes.get("BC ICMS ST", "0")),
            "BC ICMS ST Depois":    fmt_br(depois["BC ICMS ST"]),
            "Vlr ICMS ST Antes":    fmt_br(antes.get("Vlr ICMS ST", "0")),
            "Vlr ICMS ST Depois":   fmt_br(depois["Vlr ICMS ST"]),
            "Diff Vlr ICMS ST":     fmt_br(_fv(depois["Vlr ICMS ST"]) - _fv(antes.get("Vlr ICMS ST","0"))),
            # IPI
            "CST IPI Antes":        antes.get("CST IPI", ""),
            "CST IPI Depois":       depois["CST IPI"],
            "BC IPI Antes":         fmt_br(antes.get("BC IPI", "0")),
            "BC IPI Depois":        fmt_br(depois["BC IPI"]),
            "% IPI":                fmt_br(depois["% IPI"]),
            "Vlr IPI Antes":        fmt_br(antes.get("Vlr IPI", "0")),
            "Vlr IPI Depois":       fmt_br(depois["Vlr IPI"]),
            "Diff Vlr IPI":         fmt_br(_fv(depois["Vlr IPI"]) - _fv(antes.get("Vlr IPI","0"))),
            # impostoDevol
            "pDevol Antes":         fmt_br(antes.get("pDevol", "0")),
            "pDevol Depois":        fmt_br(depois["pDevol"]),
            "vIPIDevol Antes":      fmt_br(antes.get("vIPIDevol", "0")),
            "vIPIDevol Depois":     fmt_br(depois["vIPIDevol"]),
            "Diff vIPIDevol":       fmt_br(_fv(depois["vIPIDevol"]) - _fv(antes.get("vIPIDevol","0"))),
            # PIS
            "CST PIS Antes":        antes.get("CST PIS", ""),
            "CST PIS Depois":       depois["CST PIS"],
            "BC PIS Antes":         fmt_br(antes.get("BC PIS", "0")),
            "BC PIS Depois":        fmt_br(depois["BC PIS"]),
            "% PIS":                fmt_br(depois["% PIS"]),
            "Vlr PIS Antes":        fmt_br(antes.get("Vlr PIS", "0")),
            "Vlr PIS Depois":       fmt_br(depois["Vlr PIS"]),
            "Diff Vlr PIS":         fmt_br(_fv(depois["Vlr PIS"]) - _fv(antes.get("Vlr PIS","0"))),
            # COFINS
            "CST COFINS Antes":     antes.get("CST COFINS", ""),
            "CST COFINS Depois":    depois["CST COFINS"],
            "BC COFINS Antes":      fmt_br(antes.get("BC COFINS", "0")),
            "BC COFINS Depois":     fmt_br(depois["BC COFINS"]),
            "% COFINS":             fmt_br(depois["% COFINS"]),
            "Vlr COFINS Antes":     fmt_br(antes.get("Vlr COFINS", "0")),
            "Vlr COFINS Depois":    fmt_br(depois["Vlr COFINS"]),
            "Diff Vlr COFINS":      fmt_br(_fv(depois["Vlr COFINS"]) - _fv(antes.get("Vlr COFINS","0"))),
            "Tem Diferença":        "SIM" if tem_diff else "NÃO",
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
        "vBC","vICMS","vBCST","vST","vIPI","vIPIDevol","vPIS","vCOFINS","vProd"
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
                except Exception: pass

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
                                except Exception: pass
                        break

            ipi_trib = find(imposto, "IPI", "IPITrib")
            if ipi_trib:
                el = find(ipi_trib, "vIPI")
                if el is not None and el.text:
                    try: totais["vIPI"] += float(el.text)
                    except Exception: pass

            pis_pai = find(imposto, "PIS")
            if pis_pai:
                for pf in pis_pai:
                    el = find(pf, "vPIS")
                    if el is not None and el.text:
                        try: totais["vPIS"] += float(el.text)
                        except Exception: pass
                    break

            cof_pai = find(imposto, "COFINS")
            if cof_pai:
                for cf in cof_pai:
                    el = find(cf, "vCOFINS")
                    if el is not None and el.text:
                        try: totais["vCOFINS"] += float(el.text)
                        except Exception: pass
                    break

        # vIPIDevol já zerado → soma 0.00
        imp_devol = find(filho, "impostoDevol")
        if imp_devol:
            el = find(imp_devol, "IPI", "vIPIDevol")
            if el is not None and el.text:
                try: totais["vIPIDevol"] += float(el.text)
                except Exception: pass

    icms_tot = find(inf_nfe, "total", "ICMSTot")
    if icms_tot is None:
        return

    for tn, tv in {
        "vBC":       fmt(totais["vBC"]),
        "vICMS":     fmt(totais["vICMS"]),
        "vBCST":     fmt(totais["vBCST"]),
        "vST":       fmt(totais["vST"]),
        "vIPI":      fmt(totais["vIPI"]),
        "vIPIDevol": fmt(totais["vIPIDevol"]),
        "vPIS":      fmt(totais["vPIS"]),
        "vCOFINS":   fmt(totais["vCOFINS"]),
        "vProd":     fmt(totais["vProd"]),
        "vNF":       fmt(totais["vProd"] + totais["vST"] + totais["vIPI"]),
    }.items():
        el = find(icms_tot, tn)
        if el is not None:
            el.text = tv

# ─────────────────────────────────────────────
# EXCEL DE CONFERÊNCIA
# ─────────────────────────────────────────────
def gerar_excel_conferencia(todas_diferencas: list) -> bytes:
    if not todas_diferencas:
        return b""

    df  = pd.DataFrame(todas_diferencas)
    buf = io.BytesIO()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Conferência")
        ws = writer.sheets["Conferência"]

        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        COR_LARANJA  = "FF8000"
        COR_CINZA    = "444444"
        COR_BRANCO   = "FFFFFF"
        COR_VERDE_BG = "E2EFDA"
        COR_VERDE_FT = "375623"
        COR_VERM_BG  = "FCE4D6"
        COR_VERM_FT  = "843C0C"
        COR_PAR      = "F9F9F9"

        thin  = Side(style="thin",   color="CCCCCC")
        thick = Side(style="medium", color=COR_LARANJA)
        b_hdr = Border(left=thick, right=thick, top=thick, bottom=thick)
        b_nrm = Border(left=thin,  right=thin,  top=thin,  bottom=thin)

        grupos = {
            "id":      {"cols": ["Arquivo","Chave NF-e","Nro Documento","Data Emissão",
                                 "Data Entrada","Razão Social","CNPJ/CPF","UF","nItem",
                                 "Cód Item","Desc Item","NCM","CFOP XML","CFOP XLSX",
                                 "Vlr Documento"], "cor": "2E4057"},
            "icms":    {"cols": ["CST ICMS Antes","CST ICMS Depois","BC ICMS Antes",
                                 "BC ICMS Depois","% ICMS","Vlr ICMS Antes","Vlr ICMS Depois",
                                 "Diff Vlr ICMS","BC ICMS ST Antes","BC ICMS ST Depois",
                                 "Vlr ICMS ST Antes","Vlr ICMS ST Depois","Diff Vlr ICMS ST"],
                        "cor": "1F4E79"},
            "ipi":     {"cols": ["CST IPI Antes","CST IPI Depois","BC IPI Antes",
                                 "BC IPI Depois","% IPI","Vlr IPI Antes","Vlr IPI Depois",
                                 "Diff Vlr IPI"], "cor": "375623"},
            "devol":   {"cols": ["pDevol Antes","pDevol Depois",
                                 "vIPIDevol Antes","vIPIDevol Depois","Diff vIPIDevol"],
                        "cor": "7B3F00"},
            "pis":     {"cols": ["CST PIS Antes","CST PIS Depois","BC PIS Antes",
                                 "BC PIS Depois","% PIS","Vlr PIS Antes","Vlr PIS Depois",
                                 "Diff Vlr PIS"], "cor": "7B2C2C"},
            "cofins":  {"cols": ["CST COFINS Antes","CST COFINS Depois","BC COFINS Antes",
                                 "BC COFINS Depois","% COFINS","Vlr COFINS Antes",
                                 "Vlr COFINS Depois","Diff Vlr COFINS"], "cor": "843C0C"},
            "flag":    {"cols": ["Tem Diferença"], "cor": COR_LARANJA},
        }

        col_names = list(df.columns)
        col_grupo = {}
        for grp, info in grupos.items():
            for c in info["cols"]:
                if c in col_names:
                    col_grupo[col_names.index(c) + 1] = info["cor"]

        for cell in ws[1]:
            cor = col_grupo.get(cell.column, COR_CINZA)
            cell.fill      = PatternFill("solid", fgColor=cor)
            cell.font      = Font(color=COR_BRANCO, bold=True, size=9, name="Segoe UI")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = b_hdr
        ws.row_dimensions[1].height = 36

        diff_cols = [i+1 for i, c in enumerate(col_names) if c.startswith("Diff ")]
        flag_col  = col_names.index("Tem Diferença") + 1 if "Tem Diferença" in col_names else None
        fill_par  = PatternFill("solid", fgColor=COR_PAR)

        def _fv_br(v):
            try: return float(str(v).replace(",", "."))
            except Exception: return 0.0

        for row_idx, row in enumerate(
            ws.iter_rows(min_row=2, max_row=ws.max_row), start=2
        ):
            is_par = (row_idx % 2 == 0)
            for cell in row:
                cell.border    = b_nrm
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font      = Font(size=9, name="Segoe UI")
                if is_par:
                    cell.fill = fill_par

            for col_idx in diff_cols:
                cell = row[col_idx - 1]
                try:
                    v = _fv_br(cell.value or "0")
                    if v > 0:
                        cell.fill = PatternFill("solid", fgColor=COR_VERDE_BG)
                        cell.font = Font(color=COR_VERDE_FT, bold=True, size=9, name="Segoe UI")
                    elif v < 0:
                        cell.fill = PatternFill("solid", fgColor=COR_VERM_BG)
                        cell.font = Font(color=COR_VERM_FT, bold=True, size=9, name="Segoe UI")
                except Exception:
                    pass

            if flag_col:
                cell = row[flag_col - 1]
                if str(cell.value) == "SIM":
                    cell.fill = PatternFill("solid", fgColor="FFE0B2")
                    cell.font = Font(color=COR_LARANJA, bold=True, size=9, name="Segoe UI")
                else:
                    cell.fill = PatternFill("solid", fgColor=COR_VERDE_BG)
                    cell.font = Font(color=COR_VERDE_FT, bold=True, size=9, name="Segoe UI")

        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len    = len(str(ws.cell(1, col_idx).value or ""))
            for r in range(2, min(ws.max_row + 1, 200)):
                v = str(ws.cell(r, col_idx).value or "")
                max_len = max(max_len, len(v))
            ws.column_dimensions[col_letter].width = min(max_len + 3, 35)

        for cn in ["Desc Item", "Razão Social", "Arquivo", "Chave NF-e"]:
            if cn in col_names:
                ws.column_dimensions[
                    get_column_letter(col_names.index(cn) + 1)
                ].width = 40

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────
# RESULTADOS PERSISTENTES
# ─────────────────────────────────────────────
def render_resultados():
    m = st.session_state.resultado_metricas
    if m is None:
        return

    st.divider()
    st.markdown("### ✅ Resultado do Processamento")

    col_np, _ = st.columns([2, 6])
    with col_np:
        if st.button(
            "🔁 Iniciar Novo Processo",
            key="btn_novo_processo",
            use_container_width=True,
            type="primary",
        ):
            _limpar_resultados()
            st.rerun()

    st.markdown("")

    col_x, col_y, col_z, col_w = st.columns(4)
    col_x.metric("✅ Alterados",     m["ok"])
    col_y.metric("ℹ️ Sem alteração", m["info"])
    col_z.metric("❌ Erros",         m["erro"])
    col_w.metric("📋 Itens c/ diff", m["diff"])

    st.divider()

    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        excel_bytes = st.session_state.resultado_excel_bytes
        diferencas  = st.session_state.resultado_diferencas
        if excel_bytes and len(excel_bytes) > 0:
            diff_c = sum(1 for r in diferencas if r.get("Tem Diferença") == "SIM")
            st.download_button(
                label=f"📥 Excel de Conferência ({len(diferencas)} itens / {diff_c} com diff)",
                data=excel_bytes,
                file_name="conferencia_nfe_dni.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_excel_resultado",
            )
        else:
            st.info("ℹ️ Nenhum item processado para conferência.")

    with col_dl2:
        xmls = st.session_state.resultado_xmls
        if xmls:
            if len(xmls) == 1:
                nome_arq, conteudo_arq = list(xmls.items())[0]
                st.download_button(
                    label=f"⬇ Baixar XML: {nome_arq}",
                    data=conteudo_arq,
                    file_name=nome_arq,
                    mime="application/xml",
                    use_container_width=True,
                    key="dl_xml_unico_resultado",
                )
            else:
                buf = io.BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for nome_arq, conteudo_arq in xmls.items():
                        zf.writestr(nome_arq, conteudo_arq)
                buf.seek(0)
                st.download_button(
                    label=f"⬇ Baixar {len(xmls)} XMLs (ZIP)",
                    data=buf.read(),
                    file_name="xmls_modificados_dni.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="dl_zip_resultado",
                )
        else:
            st.warning("⚠️ Nenhum XML foi modificado.")

    with st.expander("📋 Log detalhado de resultados", expanded=False):
        for nome_arq, msg, status in st.session_state.resultado_log:
            icon = "✅" if status == "ok" else ("ℹ️" if status == "info" else "❌")
            st.markdown(f"{icon} `{nome_arq}` — {msg}")

# ─────────────────────────────────────────────
# INTERFACE PRINCIPAL
# ─────────────────────────────────────────────
def main():
    _init_state()
    render_sidebar()

    cfops_ativas: set[str] = {str(c) for c in st.session_state.get("cfop_ativos", ATIVOS_PADRAO)}

    st.markdown("""
    <div class="main-header">
        <h2>🧾 Enriquecedor de NF-e — DNI</h2>
        <p>Thomson Reuters · Domínio Sistemas · Processamento Fiscal Automatizado · v7.3</p>
    </div>
    """, unsafe_allow_html=True)

    if st.session_state.processamento_concluido:
        st.info(
            "📌 Resultado disponível abaixo. Clique em **🔁 Iniciar Novo Processo** para processar novos arquivos."
        )
        render_resultados()
        st.markdown("""
        <div class="footer">
            Thomson Reuters · Domínio Sistemas · Enriquecedor NF-e v7.3 · DNI
        </div>
        """, unsafe_allow_html=True)
        return

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        st.markdown("#### 📂 Arquivo XLSX — Domínio Sistemas")
        arquivo_xlsx = st.file_uploader(
            "xlsx", type=["xlsx"], key="xlsx", label_visibility="collapsed"
        )
    with col_up2:
        st.markdown("#### 📄 Arquivos XML de NF-e")
        arquivos_xml = st.file_uploader(
            "xmls", type=["xml", "zip"], accept_multiple_files=True,
            key="xmls", label_visibility="collapsed"
        )

    lista_ativos_str = sorted(cfops_ativas)
    st.info(
        f"🔢 **CFOPs filtrados (lidos do XML):** "
        f"{', '.join(lista_ativos_str) if lista_ativos_str else '⚠️ Nenhum CFOP ativo — configure na barra lateral.'}"
    )

    st.divider()

    if st.button("▶ Processar XMLs", type="primary", key="btn_processar", use_container_width=True):
        if not arquivo_xlsx:
            st.error("❌ Selecione o arquivo XLSX.")
            st.stop()
        if not arquivos_xml:
            st.error("❌ Selecione ao menos um XML ou ZIP.")
            st.stop()
        if not cfops_ativas:
            st.error("❌ Nenhum CFOP ativo. Configure na barra lateral esquerda.")
            st.stop()

        with st.spinner("🔄 Lendo XLSX..."):
            dados_indexados = ler_xlsx(arquivo_xlsx.read())

        if not dados_indexados:
            st.error("❌ Nenhum item válido no XLSX.")
            st.stop()

        col_a, col_b, col_c = st.columns(3)
        col_a.metric("📊 Itens indexados", len(dados_indexados))
        col_b.metric("🔢 CFOPs filtro",    len(cfops_ativas))
        col_c.metric("📁 Arquivos",        len(arquivos_xml))

        with st.expander("🔍 Debug — primeiros 5 itens indexados"):
            for i, ((ch, seq), d) in enumerate(list(dados_indexados.items())[:5]):
                st.code(
                    f"Chave: {ch} | Seq: {seq} | CFOP XLSX: {d['cfop']}\n"
                    f"CST ICMS: {d['cst_icms']} | vICMS: {d['vlr_icms']} | "
                    f"BC ST: {d['base_icms_st']} | vST: {d['vlr_icms_st']}\n"
                    f"BC IPI: {d['base_ipi']} | % IPI: {d['perc_ipi']} | "
                    f"vIPI: {d['vlr_ipi']} | tem_ipi: {d['tem_ipi']}\n"
                    f"BC PIS/COF: {d['bc_pis_cofins']}"
                )

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

        excel_bytes = gerar_excel_conferencia(todas_diferencas) if todas_diferencas else b""

        st.session_state.resultado_excel_bytes  = excel_bytes
        st.session_state.resultado_xmls         = xmls_modificados
        st.session_state.resultado_log          = resultados
        st.session_state.resultado_diferencas   = todas_diferencas
        st.session_state.resultado_metricas     = {
            "ok":   sum(1 for _, _, s in resultados if s == "ok"),
            "info": sum(1 for _, _, s in resultados if s == "info"),
            "erro": sum(1 for _, _, s in resultados if s == "erro"),
            "diff": sum(1 for r in todas_diferencas if r.get("Tem Diferença") == "SIM"),
        }
        st.session_state.processamento_concluido = True
        st.rerun()

    st.markdown("""
    <div class="footer">
        Thomson Reuters · Domínio Sistemas · Enriquecedor NF-e v7.3 · DNI
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
