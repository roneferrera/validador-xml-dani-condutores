"""
DNI – Alteração de XMLs NF-e
Versão Streamlit 1.0
"""

import io
import re
import zipfile
import xml.etree.ElementTree as ET
from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict

import streamlit as st

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────

NS  = "http://www.portalfiscal.inf.br/nfe"
NSP = f"{{{NS}}}"
ET.register_namespace("", NS)

CFOPS_PADRAO = {
    "2201","2949","1411","1202","1410",
    "1201","1949","2411","2202","2410","2603"
}

GRUPOS_ST = ["ICMS10","ICMS30","ICMS70","ICMS90",
             "CSOSN201","CSOSN202","CSOSN203","CSOSN900"]

ORDEM_TAGS_ICMS = [
    "orig","CST","CSOSN","modBC","vBC","pICMS","vICMS",
    "modBCST","pMVAST","pRedBCCST","vBCST","pICMSST","vICMSST",
    "vICMSSTDeson","vBCSTDeson"
]

# ─────────────────────────────────────────────
# FUNÇÕES AUXILIARES
# ─────────────────────────────────────────────

def arredondar(valor: float, casas: int = 2) -> Decimal:
    return Decimal(str(valor)).quantize(
        Decimal("0." + "0" * casas), rounding=ROUND_HALF_UP
    )

def fmt(valor: Decimal, casas: int = 2) -> str:
    return str(valor.quantize(Decimal("0." + "0" * casas), rounding=ROUND_HALF_UP))

def limpar_valor(texto) -> float:
    if not texto or str(texto).strip().lower() in ("nan", "", "none"):
        return 0.0
    texto = str(texto).strip().replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0

def limpar_chave(chave: str) -> str:
    return str(chave).strip().lstrip("'").strip()

def t(nome: str) -> str:
    return f"{NSP}{nome}"

def dec(texto) -> Decimal:
    try:
        return Decimal(str(texto or "0").strip())
    except:
        return Decimal("0")

# ─────────────────────────────────────────────
# LEITURA DO CSV
# ─────────────────────────────────────────────

def ler_csv(conteudo: str):
    todos = []
    indexado = {}

    for linha in conteudo.split("\n"):
        linha = linha.strip()
        if not linha:
            continue

        registro = {}
        partes = re.split(r",\s*(?=[A-Za-záéíóúàâêôãõüçÁÉÍÓÚÀÂÊÔÃÕÜÇ0-9\s\.\/]+:)", linha)
        for parte in partes:
            if ":" in parte:
                k, _, v = parte.partition(":")
                registro[k.strip()] = v.strip()

        if not registro:
            continue

        todos.append(registro)

        chave_nfe = limpar_chave(registro.get("Chave Nfe/Cte", ""))
        try:
            seq = int(float(registro.get("Sequencia", "0")))
        except:
            seq = 0

        if not chave_nfe or seq == 0:
            continue

        cfop = registro.get("Cfop", "").strip()

        indexado[(chave_nfe, seq)] = {
            "cfop":          cfop,
            "cod_item":      registro.get("Cod Item", "").strip(),
            "desc_item":     registro.get("Desc Item", "").strip(),
            "ncm":           registro.get("NCM", "").strip(),
            "nro_documento": registro.get("Nro Documento", "").strip(),
            "razao_social":  registro.get("Razao Social", "").strip(),
            "cnpj":          registro.get("CNPJ-CPF", "").strip(),
            "data_entrada":  registro.get("Data Entrada", "").strip(),
            "data_emissao":  registro.get("Data Emissao", "").strip(),
            "uf":            registro.get("UF", "").strip(),
            "vlr_documento": limpar_valor(registro.get("Vlr Documento", "0")),
            "vlr_produto":   limpar_valor(registro.get("Vlr Produto", "0")),
            "cst_icms":      registro.get("CST ICMS", "").strip(),
            "base_icms":     limpar_valor(registro.get("Base Icms", "0")),
            "perc_icms":     limpar_valor(registro.get("Perc ICms", "0")),
            "vlr_icms":      limpar_valor(registro.get("Vlr Icms", "0")),
            "base_icms_st":  limpar_valor(registro.get("Base Icms St", "0")),
            "vlr_icms_st":   limpar_valor(registro.get("Vlr Icms St", "0")),
            "cst_ipi":       registro.get("CST IPI", "").strip(),
            "base_ipi":      limpar_valor(registro.get("Base Ipi", "0")),
            "perc_ipi":      limpar_valor(registro.get("Perc Ipi", "0")),
            "vlr_ipi":       limpar_valor(registro.get("Vlr Ipi", "0")),
            "perc_pis":      limpar_valor(registro.get("Perc Pis", "0")),
            "vlr_pis":       limpar_valor(registro.get("Vlr Pis", "0")),
            "perc_cofins":   limpar_valor(registro.get("Perc Cofins", "0")),
            "vlr_cofins":    limpar_valor(registro.get("Vlr Cofins", "0")),
            "chave_nfe":     chave_nfe,
            "sequencia":     seq,
        }

    return todos, indexado

# ─────────────────────────────────────────────
# PROCESSAMENTO ICMS ST
# ─────────────────────────────────────────────

def determinar_modBCST(grupo_el) -> str:
    mod_el  = grupo_el.find(t("modBCST"))
    pmva_el = grupo_el.find(t("pMVAST"))
    if mod_el is not None and mod_el.text and mod_el.text.strip() in {"0","1","2","3","4","5","6"}:
        return mod_el.text.strip()
    if pmva_el is not None:
        try:
            if float((pmva_el.text or "0").replace(",",".")) > 0:
                return "4"
        except:
            pass
    return "4"

def aplicar_modBCST(grupo_el, mod_valor: str):
    mod_el = grupo_el.find(t("modBCST"))
    if mod_el is None:
        mod_el = ET.SubElement(grupo_el, t("modBCST"))
    mod_el.text = mod_valor

    pmva_el = grupo_el.find(t("pMVAST"))
    if mod_valor == "4":
        if pmva_el is None:
            pmva_el = ET.SubElement(grupo_el, t("pMVAST"))
            pmva_el.text = "0.0000"
        elif not pmva_el.text or pmva_el.text.strip() in ("", "nan"):
            pmva_el.text = "0.0000"
    else:
        if pmva_el is not None:
            grupo_el.remove(pmva_el)

    filhos = list(grupo_el)
    grupo_el[:] = sorted(
        filhos,
        key=lambda el: ORDEM_TAGS_ICMS.index(ET.QName(el.tag).localname)
        if ET.QName(el.tag).localname in ORDEM_TAGS_ICMS else 999
    )

# ─────────────────────────────────────────────
# PROCESSAMENTO DE ITEM
# ─────────────────────────────────────────────

def processar_item(det, dados_item: dict, nitem: int) -> dict:
    resultado = {
        "antes_vBCST": Decimal("0"), "antes_vICMSST": Decimal("0"),
        "antes_vIPI": Decimal("0"),  "antes_vIPIDevol": Decimal("0"),
        "antes_vBC_pis": Decimal("0"), "antes_vPIS": Decimal("0"),
        "antes_vBC_cofins": Decimal("0"), "antes_vCOFINS": Decimal("0"),
        "vBCST": Decimal("0"), "vICMSST": Decimal("0"),
        "vIPI": Decimal("0"),  "vIPIDevol": Decimal("0"),
        "vPIS": Decimal("0"),  "vCOFINS": Decimal("0"),
        "vBC_pis": Decimal("0"), "vBC_cofins": Decimal("0"),
        "alterado": False, "modBCST": "4",
    }

    imposto = det.find(t("imposto"))
    if imposto is None:
        return resultado

    vBCST_novo   = arredondar(dados_item["base_icms_st"])
    vICMSST_novo = arredondar(dados_item["vlr_icms_st"])
    vIPI_novo    = arredondar(dados_item["vlr_ipi"])
    vICMS_novo   = arredondar(dados_item["vlr_icms"])

    # ICMS ST
    icms_el = imposto.find(t("ICMS"))
    if icms_el is not None:
        for nome_grupo in GRUPOS_ST:
            grupo = icms_el.find(t(nome_grupo))
            if grupo is None:
                continue
            vbcst_a   = grupo.find(t("vBCST"))
            vicmsst_a = grupo.find(t("vICMSST"))
            resultado["antes_vBCST"]   = dec(vbcst_a.text   if vbcst_a   is not None else "0")
            resultado["antes_vICMSST"] = dec(vicmsst_a.text if vicmsst_a is not None else "0")

            mod_valor = determinar_modBCST(grupo)
            aplicar_modBCST(grupo, mod_valor)
            resultado["modBCST"] = mod_valor

            if vbcst_a is not None:
                vbcst_a.text = fmt(vBCST_novo)
            if vicmsst_a is not None:
                vicmsst_a.text = fmt(vICMSST_novo)

            resultado["alterado"] = True
            break

    resultado["vBCST"]   = vBCST_novo
    resultado["vICMSST"] = vICMSST_novo

    # IPI
    ipi_el = imposto.find(t("IPI"))
    if ipi_el is not None:
        ipi_trib = ipi_el.find(t("IPITrib"))
        if ipi_trib is not None:
            vipi_el = ipi_trib.find(t("vIPI"))
            if vipi_el is not None:
                resultado["antes_vIPI"] = dec(vipi_el.text)
                vipi_el.text = fmt(vIPI_novo)
                resultado["alterado"] = True

    resultado["vIPI"] = vIPI_novo

    imp_devol = det.find(t("impostoDevol"))
    if imp_devol is not None:
        vipi_devol_el = imp_devol.find(f".//{t('vIPIDevol')}")
        if vipi_devol_el is not None:
            resultado["antes_vIPIDevol"] = dec(vipi_devol_el.text)
            vipi_devol_el.text = fmt(vIPI_novo)
            resultado["vIPIDevol"] = vIPI_novo

    # PIS e COFINS
    vDoc    = arredondar(dados_item["vlr_documento"])
    pPIS    = arredondar(dados_item["perc_pis"])
    pCOFINS = arredondar(dados_item["perc_cofins"])

    vBC_novo = max(Decimal("0"), vDoc - vICMS_novo - vICMSST_novo - vIPI_novo)
    vBC_novo = arredondar(float(vBC_novo))

    vPIS_novo    = arredondar(float(vBC_novo) * float(pPIS)    / 100)
    vCOFINS_novo = arredondar(float(vBC_novo) * float(pCOFINS) / 100)

    pis_el = imposto.find(t("PIS"))
    if pis_el is not None:
        for tipo in ["PISAliq","PISQtde","PISNT","PISOutr"]:
            grp = pis_el.find(t(tipo))
            if grp is not None:
                el_bc  = grp.find(t("vBC"))
                el_pis = grp.find(t("vPIS"))
                if el_bc  is not None:
                    resultado["antes_vBC_pis"] = dec(el_bc.text)
                    el_bc.text = fmt(vBC_novo)
                if el_pis is not None:
                    resultado["antes_vPIS"] = dec(el_pis.text)
                    el_pis.text = fmt(vPIS_novo)
                resultado["alterado"] = True
                break

    cofins_el = imposto.find(t("COFINS"))
    if cofins_el is not None:
        for tipo in ["COFINSAliq","COFINSQtde","COFINSNT","COFINSOutr"]:
            grp = cofins_el.find(t(tipo))
            if grp is not None:
                el_bc     = grp.find(t("vBC"))
                el_cofins = grp.find(t("vCOFINS"))
                if el_bc     is not None:
                    resultado["antes_vBC_cofins"] = dec(el_bc.text)
                    el_bc.text = fmt(vBC_novo)
                if el_cofins is not None:
                    resultado["antes_vCOFINS"] = dec(el_cofins.text)
                    el_cofins.text = fmt(vCOFINS_novo)
                resultado["alterado"] = True
                break

    resultado["vPIS"]       = vPIS_novo
    resultado["vCOFINS"]    = vCOFINS_novo
    resultado["vBC_pis"]    = vBC_novo
    resultado["vBC_cofins"] = vBC_novo
    return resultado

def capturar_originais(det) -> dict:
    vals = {k: Decimal("0") for k in ["vBCST","vICMSST","vIPI","vIPIDevol","vPIS","vCOFINS"]}
    imposto = det.find(t("imposto"))
    if imposto is None:
        return vals
    icms_el = imposto.find(t("ICMS"))
    if icms_el is not None:
        for g in GRUPOS_ST:
            grupo = icms_el.find(t(g))
            if grupo is not None:
                for campo, chave in [("vBCST","vBCST"),("vICMSST","vICMSST")]:
                    el = grupo.find(t(campo))
                    if el is not None:
                        try: vals[chave] = Decimal(el.text or "0")
                        except: pass
                break
    ipi_el = imposto.find(t("IPI"))
    if ipi_el is not None:
        el = ipi_el.find(f".//{t('vIPI')}")
        if el is not None:
            try: vals["vIPI"] = Decimal(el.text or "0")
            except: pass
    imp_devol = det.find(t("impostoDevol"))
    if imp_devol is not None:
        el = imp_devol.find(f".//{t('vIPIDevol')}")
        if el is not None:
            try: vals["vIPIDevol"] = Decimal(el.text or "0")
            except: pass
    pis_el = imposto.find(t("PIS"))
    if pis_el is not None:
        el = pis_el.find(f".//{t('vPIS')}")
        if el is not None:
            try: vals["vPIS"] = Decimal(el.text or "0")
            except: pass
    cofins_el = imposto.find(t("COFINS"))
    if cofins_el is not None:
        el = cofins_el.find(f".//{t('vCOFINS')}")
        if el is not None:
            try: vals["vCOFINS"] = Decimal(el.text or "0")
            except: pass
    return vals

def recalcular_totais(raiz, acumulados: dict):
    icms_tot = raiz.find(f".//{t('ICMSTot')}")
    if icms_tot is None:
        return
    def upd(nome_tag, valor):
        el = icms_tot.find(t(nome_tag))
        if el is not None:
            el.text = fmt(valor)
    upd("vBCST",     acumulados["vBCST"])
    upd("vST",       acumulados["vICMSST"])
    upd("vIPI",      acumulados["vIPI"])
    upd("vIPIDevol", acumulados["vIPIDevol"])
    upd("vPIS",      acumulados["vPIS"])
    upd("vCOFINS",   acumulados["vCOFINS"])

# ─────────────────────────────────────────────
# PROCESSAMENTO DO XML
# ─────────────────────────────────────────────

def processar_xml_bytes(nome: str, conteudo_bytes: bytes,
                        dados_csv: dict, cfops_validos: set):
    logs = []
    registros_excel = []

    try:
        raiz = ET.fromstring(conteudo_bytes)
    except ET.ParseError as e:
        return None, [], [f"❌ {nome}: erro de parse XML — {e}"]

    inf_nfe = raiz.find(f".//{t('infNFe')}")
    if inf_nfe is None:
        return None, [], [f"⚠️ {nome}: <infNFe> não encontrado"]

    chave_xml = inf_nfe.get("Id", "").replace("NFe", "").strip()
    if not chave_xml:
        return None, [], [f"⚠️ {nome}: chave NFe vazia"]

    itens_csv = {
        seq: d for (ch, seq), d in dados_csv.items()
        if ch == chave_xml and d["cfop"] in cfops_validos
    }

    if not itens_csv:
        return None, [], [f"ℹ️ {nome}: nenhum item com CFOP válido — não alterado"]

    dets = raiz.findall(f".//{t('det')}")
    acumulados = {k: Decimal("0") for k in ["vBCST","vICMSST","vIPI","vIPIDevol","vPIS","vCOFINS"]}
    itens_alterados = 0

    for det in dets:
        nitem = int(det.get("nItem", "0"))
        if nitem not in itens_csv:
            orig = capturar_originais(det)
            for k in acumulados:
                acumulados[k] += orig[k]
            continue

        dados_item = itens_csv[nitem]
        resultado  = processar_item(det, dados_item, nitem)

        if resultado["alterado"]:
            itens_alterados += 1
            for k in ["vBCST","vICMSST","vIPI","vIPIDevol","vPIS","vCOFINS"]:
                acumulados[k] += resultado[k]

            registros_excel.append({
                "chave_nfe":        chave_xml,
                "nro_documento":    dados_item["nro_documento"],
                "razao_social":     dados_item["razao_social"],
                "cnpj":             dados_item["cnpj"],
                "uf":               dados_item["uf"],
                "data_entrada":     dados_item["data_entrada"],
                "sequencia":        nitem,
                "cfop":             dados_item["cfop"],
                "cod_item":         dados_item["cod_item"],
                "desc_item":        dados_item["desc_item"],
                "vlr_documento":    dados_item["vlr_documento"],
                "base_icms_st_csv": dados_item["base_icms_st"],
                "vlr_icms_st_csv":  dados_item["vlr_icms_st"],
                "antes_base_st":    float(resultado["antes_vBCST"]),
                "antes_vlr_st":     float(resultado["antes_vICMSST"]),
                "vlr_ipi_csv":      dados_item["vlr_ipi"],
                "antes_vlr_ipi":    float(resultado["antes_vIPI"]),
                "vlr_pis_novo":     float(resultado["vPIS"]),
                "antes_vlr_pis":    float(resultado["antes_vPIS"]),
                "vlr_cofins_novo":  float(resultado["vCOFINS"]),
                "antes_vlr_cofins": float(resultado["antes_vCOFINS"]),
            })
            logs.append(f"✅ {nome} | nItem={nitem} | {dados_item['cod_item']} | ST: {float(resultado['antes_vICMSST']):.2f}→{dados_item['vlr_icms_st']:.2f} | IPI: {float(resultado['antes_vIPI']):.2f}→{dados_item['vlr_ipi']:.2f}")
        else:
            orig = capturar_originais(det)
            for k in acumulados:
                acumulados[k] += orig[k]

    if itens_alterados == 0:
        return None, [], [f"ℹ️ {nome}: itens encontrados mas nenhuma alteração aplicada"]

    recalcular_totais(raiz, acumulados)

    buf = io.BytesIO()
    tree = ET.ElementTree(raiz)
    tree.write(buf, encoding="UTF-8", xml_declaration=True)
    buf.seek(0)

    logs.insert(0, f"✅ {nome}: {itens_alterados} item(ns) alterado(s)")
    return buf.getvalue(), registros_excel, logs

# ─────────────────────────────────────────────
# GERAÇÃO DO EXCEL
# ─────────────────────────────────────────────

def fill(cor):
    return PatternFill(start_color=cor, end_color=cor, fill_type="solid")

def borda():
    s = Side(style="thin", color="B8CCE4")
    return Border(left=s, right=s, top=s, bottom=s)

def gerar_excel_bytes(alterados: list) -> bytes:
    wb = Workbook()

    # ── Aba Resumo por Documento ──────────────────────────────────────
    ws = wb.active
    ws.title = "Resumo por Documento"

    cabecalhos = [
        "Nro Documento","Razão Social","CNPJ/CPF","UF",
        "Data Entrada","Qtd Itens",
        "Σ Base ST (CSV)","Σ Vlr ST (CSV)",
        "Σ Vlr IPI (CSV)","Σ Vlr PIS (novo)","Σ Vlr COFINS (novo)",
        "Chave NFe",
    ]

    # Título
    ws.merge_cells(f"A1:{get_column_letter(len(cabecalhos))}1")
    c = ws["A1"]
    c.value = "DNI – RESUMO DE ALTERAÇÕES POR DOCUMENTO"
    c.fill  = fill("003366")
    c.font  = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # Cabeçalhos
    for col_idx, cab in enumerate(cabecalhos, 1):
        c = ws.cell(row=2, column=col_idx, value=cab)
        c.fill  = fill("004C97")
        c.font  = Font(bold=True, color="FFFFFF", size=8, name="Calibri")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borda()
    ws.row_dimensions[2].height = 28

    # Agrupa por documento
    docs = defaultdict(lambda: {
        "razao_social":"","cnpj":"","uf":"","data_entrada":"",
        "chave_nfe":"","qtd":0,
        "base_st":0,"vlr_st":0,"vlr_ipi":0,"vlr_pis":0,"vlr_cofins":0,
    })
    for reg in alterados:
        d = docs[reg["nro_documento"]]
        d["razao_social"] = reg["razao_social"]
        d["cnpj"]         = reg["cnpj"]
        d["uf"]           = reg["uf"]
        d["data_entrada"] = reg["data_entrada"]
        d["chave_nfe"]    = reg["chave_nfe"]
        d["qtd"]         += 1
        d["base_st"]     += reg["base_icms_st_csv"]
        d["vlr_st"]      += reg["vlr_icms_st_csv"]
        d["vlr_ipi"]     += reg["vlr_ipi_csv"]
        d["vlr_pis"]     += reg["vlr_pis_novo"]
        d["vlr_cofins"]  += reg["vlr_cofins_novo"]

    cols_num = {7, 8, 9, 10, 11}
    for idx, (nro_doc, d) in enumerate(sorted(docs.items()), 3):
        cor = "D6E4F0" if idx % 2 == 0 else "FFFFFF"
        vals = [
            nro_doc, d["razao_social"], d["cnpj"], d["uf"],
            d["data_entrada"], d["qtd"],
            round(d["base_st"],2), round(d["vlr_st"],2),
            round(d["vlr_ipi"],2), round(d["vlr_pis"],2),
            round(d["vlr_cofins"],2), d["chave_nfe"],
        ]
        for col_idx, val in enumerate(vals, 1):
            c = ws.cell(row=idx, column=col_idx, value=val)
            c.fill   = fill(cor)
            c.font   = Font(size=8, name="Calibri")
            c.border = borda()
            c.alignment = Alignment(vertical="center")
            if col_idx in cols_num:
                c.number_format = "#,##0.00"
                c.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[idx].height = 14

    # Linha total
    linha_tot = len(docs) + 3
    c = ws.cell(row=linha_tot, column=1, value="TOTAL GERAL")
    c.fill = fill("003366")
    c.font = Font(bold=True, color="FFFFFF", size=9, name="Calibri")
    c.border = borda()
    for col_idx in range(2, len(cabecalhos)+1):
        c = ws.cell(row=linha_tot, column=col_idx)
        c.fill = fill("003366")
        c.border = borda()
        if col_idx in cols_num:
            letra = get_column_letter(col_idx)
            c.value = f"=SUM({letra}3:{letra}{linha_tot-1})"
            c.number_format = "#,##0.00"
            c.font = Font(bold=True, color="FFFFFF", size=9, name="Calibri")
            c.alignment = Alignment(horizontal="right", vertical="center")

    # Larguras
    largs = [14, 35, 18, 5, 12, 8, 13, 11, 12, 12, 14, 46]
    for i, larg in enumerate(largs, 1):
        ws.column_dimensions[get_column_letter(i)].width = larg

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(cabecalhos))}2"

    # ── Aba Conferência (detalhe por item) ───────────────────────────
    ws2 = wb.create_sheet("Conferência por Item")

    cabs2 = [
        "Nro Doc","Razão Social","CNPJ","UF","Dt Entrada","Seq","CFOP",
        "Cód Item","Desc Item","Vlr Documento",
        "Base ST (XML)","Vlr ST (XML)","Base ST (CSV)","Vlr ST (CSV)","Δ Vlr ST",
        "IPI (XML)","IPI (CSV)","Δ IPI",
        "PIS (XML)","PIS (CSV)","Δ PIS",
        "COFINS (XML)","COFINS (CSV)","Δ COFINS",
    ]
    n2 = len(cabs2)

    ws2.merge_cells(f"A1:{get_column_letter(n2)}1")
    c = ws2["A1"]
    c.value = "DNI – CONFERÊNCIA DE ALTERAÇÕES POR ITEM"
    c.fill  = fill("003366")
    c.font  = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 22

    for col_idx, cab in enumerate(cabs2, 1):
        c = ws2.cell(row=2, column=col_idx, value=cab)
        c.fill  = fill("004C97")
        c.font  = Font(bold=True, color="FFFFFF", size=8, name="Calibri")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borda()
    ws2.row_dimensions[2].height = 28

    cols_num2 = set(range(10, n2+1))
    cols_delta = {15, 18, 21, 24}

    for idx, reg in enumerate(alterados, 3):
        cor = "D6E4F0" if idx % 2 == 0 else "FFFFFF"
        delta_st     = reg["vlr_icms_st_csv"]  - reg["antes_vlr_st"]
        delta_ipi    = reg["vlr_ipi_csv"]       - reg["antes_vlr_ipi"]
        delta_pis    = reg["vlr_pis_novo"]      - reg["antes_vlr_pis"]
        delta_cofins = reg["vlr_cofins_novo"]   - reg["antes_vlr_cofins"]

        vals2 = [
            reg["nro_documento"], reg["razao_social"], reg["cnpj"], reg["uf"],
            reg["data_entrada"],  reg["sequencia"],    reg["cfop"],
            reg["cod_item"],      reg["desc_item"],    reg["vlr_documento"],
            reg["antes_base_st"], reg["antes_vlr_st"],
            reg["base_icms_st_csv"], reg["vlr_icms_st_csv"], delta_st,
            reg["antes_vlr_ipi"], reg["vlr_ipi_csv"],  delta_ipi,
            reg["antes_vlr_pis"], reg["vlr_pis_novo"], delta_pis,
            reg["antes_vlr_cofins"], reg["vlr_cofins_novo"], delta_cofins,
        ]
        for col_idx, val in enumerate(vals2, 1):
            c = ws2.cell(row=idx, column=col_idx, value=val)
            c.fill   = fill(cor)
            c.font   = Font(size=8, name="Calibri")
            c.border = borda()
            c.alignment = Alignment(vertical="center")
            if col_idx in cols_num2:
                c.number_format = "#,##0.00"
                c.alignment = Alignment(horizontal="right", vertical="center")
            if col_idx in cols_delta and isinstance(val, float) and abs(val) > 0.005:
                c.fill = fill("FFF0E0")
                c.font = Font(size=8, name="Calibri", bold=True, color="E87722")
        ws2.row_dimensions[idx].height = 14

    largs2 = [12, 30, 18, 5, 11, 6, 7, 10, 40, 12,
              12, 10, 12, 10, 9, 10, 9, 9, 10, 9, 9, 11, 11, 9]
    for i, larg in enumerate(largs2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = larg

    ws2.freeze_panes = "A3"
    ws2.auto_filter.ref = f"A2:{get_column_letter(n2)}2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─────────────────────────────────────────────
# INTERFACE STREAMLIT
# ─────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="DNI – Alteração de XMLs NF-e",
        page_icon="📄",
        layout="wide",
    )

    st.title("📄 DNI – Alteração de XMLs NF-e")
    st.caption("Recalcula ICMS ST · IPI · PIS · COFINS com base no CSV do Domínio")

    # ── Sidebar: configuração de CFOPs ────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Configuração")
        st.markdown("**CFOPs a processar**")
        cfops_txt = st.text_area(
            "Um CFOP por linha",
            value="\n".join(sorted(CFOPS_PADRAO)),
            height=220,
        )
        cfops_validos = {c.strip() for c in cfops_txt.splitlines() if c.strip().isdigit()}
        st.info(f"{len(cfops_validos)} CFOPs ativos")

        st.divider()
        st.markdown("**Como usar**")
        st.markdown(
            "1. Ajuste os CFOPs ao lado\n"
            "2. Faça upload do CSV (`es1.csv`)\n"
            "3. Faça upload dos XMLs (`.xml` individuais ou `.zip`)\n"
            "4. Clique em **Processar**\n"
            "5. Baixe os XMLs alterados e o Excel de conferência"
        )

    # ── Upload ────────────────────────────────────────────────────────
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1️⃣ CSV do Domínio")
        csv_file = st.file_uploader(
            "Selecione o arquivo CSV (es1.csv)",
            type=["csv", "txt"],
            key="csv_upload",
        )

    with col2:
        st.subheader("2️⃣ XMLs de NF-e")
        xml_files = st.file_uploader(
            "Selecione os XMLs ou um arquivo ZIP",
            type=["xml", "zip"],
            accept_multiple_files=True,
            key="xml_upload",
        )

    if not csv_file or not xml_files:
        st.info("⬆️ Faça o upload do CSV e dos XMLs para continuar.")
        return

    # ── Botão processar ───────────────────────────────────────────────
    if st.button("🚀 Processar", type="primary", use_container_width=True):

        # Lê CSV
        with st.spinner("Lendo CSV..."):
            conteudo_csv = csv_file.read().decode("utf-8", errors="replace")
            todos_csv, dados_indexados = ler_csv(conteudo_csv)

        st.success(f"CSV lido: **{len(todos_csv)}** registros | **{len(dados_indexados)}** itens indexados")

        # Coleta XMLs (suporte a .zip)
        xmls_para_processar = {}  # nome → bytes

        for f in xml_files:
            if f.name.lower().endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(f.read())) as zf:
                    for membro in zf.namelist():
                        if membro.lower().endswith(".xml"):
                            xmls_para_processar[membro.split("/")[-1]] = zf.read(membro)
            elif f.name.lower().endswith(".xml"):
                xmls_para_processar[f.name] = f.read()

        st.info(f"XMLs encontrados: **{len(xmls_para_processar)}**")

        # Processa
        todos_alterados  = []
        xmls_alterados   = {}   # nome → bytes
        todos_logs       = []
        n_alterados      = 0
        n_nao_alterados  = 0

        progress = st.progress(0, text="Processando XMLs...")
        total = len(xmls_para_processar)

        for i, (nome, conteudo_bytes) in enumerate(xmls_para_processar.items()):
            xml_saida, registros, logs = processar_xml_bytes(
                nome, conteudo_bytes, dados_indexados, cfops_validos
            )
            todos_logs.extend(logs)

            if xml_saida is not None:
                xmls_alterados[nome] = xml_saida
                todos_alterados.extend(registros)
                n_alterados += 1
            else:
                n_nao_alterados += 1

            progress.progress((i + 1) / total, text=f"Processando {i+1}/{total}...")

        progress.empty()

        # ── Resumo ────────────────────────────────────────────────────
        st.divider()
        st.subheader("📊 Resumo do Processamento")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("XMLs encontrados",  total)
        m2.metric("XMLs alterados",    n_alterados)
        m3.metric("Não alterados",     n_nao_alterados)
        m4.metric("Itens modificados", len(todos_alterados))

        if todos_alterados:
            # Tabela resumo por documento
            st.subheader("📋 Alterações por Documento")

            docs_resumo = defaultdict(lambda: {
                "Razão Social":"","CNPJ":"","UF":"",
                "Qtd Itens":0,"Σ Vlr ST":0.0,"Σ Vlr IPI":0.0,
                "Σ Vlr PIS":0.0,"Σ Vlr COFINS":0.0,
            })
            for reg in todos_alterados:
                d = docs_resumo[reg["nro_documento"]]
                d["Razão Social"] = reg["razao_social"]
                d["CNPJ"]         = reg["cnpj"]
                d["UF"]           = reg["uf"]
                d["Qtd Itens"]   += 1
                d["Σ Vlr ST"]    += reg["vlr_icms_st_csv"]
                d["Σ Vlr IPI"]   += reg["vlr_ipi_csv"]
                d["Σ Vlr PIS"]   += reg["vlr_pis_novo"]
                d["Σ Vlr COFINS"]+= reg["vlr_cofins_novo"]

            tabela = []
            for nro_doc, d in sorted(docs_resumo.items()):
                tabela.append({
                    "Nro Doc":     nro_doc,
                    "Razão Social":d["Razão Social"],
                    "UF":          d["UF"],
                    "Qtd Itens":   d["Qtd Itens"],
                    "Σ Vlr ST":    round(d["Σ Vlr ST"],    2),
                    "Σ Vlr IPI":   round(d["Σ Vlr IPI"],   2),
                    "Σ Vlr PIS":   round(d["Σ Vlr PIS"],   2),
                    "Σ Vlr COFINS":round(d["Σ Vlr COFINS"],2),
                })

            st.dataframe(tabela, use_container_width=True, hide_index=True)

        # ── Log detalhado (expansível) ────────────────────────────────
        with st.expander(f"📜 Log detalhado ({len(todos_logs)} entradas)"):
            for linha in todos_logs:
                if linha.startswith("✅"):
                    st.success(linha)
                elif linha.startswith("⚠️"):
                    st.warning(linha)
                elif linha.startswith("❌"):
                    st.error(linha)
                else:
                    st.info(linha)

        # ── Downloads ─────────────────────────────────────────────────
        st.divider()
        st.subheader("⬇️ Downloads")

        dl1, dl2 = st.columns(2)

        # ZIP com XMLs alterados
        with dl1:
            if xmls_alterados:
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for nome_xml, conteudo_xml in xmls_alterados.items():
                        zf.writestr(nome_xml, conteudo_xml)
                zip_buf.seek(0)

                st.download_button(
                    label=f"📦 Baixar XMLs alterados ({len(xmls_alterados)} arquivos)",
                    data=zip_buf.getvalue(),
                    file_name="xmls_alterados_dni.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
            else:
                st.warning("Nenhum XML foi alterado.")

        # Excel de conferência
        with dl2:
            if todos_alterados and EXCEL_OK:
                excel_bytes = gerar_excel_bytes(todos_alterados)
                st.download_button(
                    label="📊 Baixar Excel de conferência",
                    data=excel_bytes,
                    file_name="dni_conferencia.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            elif not EXCEL_OK:
                st.error("openpyxl não instalado. Execute: pip install openpyxl")
            else:
                st.warning("Nenhum item alterado para exportar.")


if __name__ == "__main__":
    main()
