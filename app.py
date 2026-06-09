import streamlit as st
import pandas as pd
import io
import os
import re
import zipfile
from lxml import etree
from copy import deepcopy

# ─────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(page_title="Enriquecedor de NF-e", layout="wide")
st.title("📄 Enriquecedor de NF-e — XLSX + XML")

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def limpar_chave(valor: str) -> str:
    if not valor:
        return ""
    v = str(valor).strip()
    v = re.sub(r"[^0-9]", "", v)
    return v

def limpar_valor(valor) -> float:
    if valor is None:
        return 0.0
    s = str(valor).strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0

def fmt(valor: float) -> str:
    """Formata float com 2 casas decimais para XML."""
    return f"{valor:.2f}"

# ─────────────────────────────────────────────
# LEITURA DO XLSX
# ─────────────────────────────────────────────
def ler_xlsx(conteudo_bytes: bytes):
    """
    Lê o XLSX do Domínio e retorna dict indexado por (chave44, sequencia_int).
    """
    todos = []
    indexado = {}

    try:
        df = pd.read_excel(io.BytesIO(conteudo_bytes), dtype=str, engine="openpyxl")
    except Exception as e:
        st.error(f"Erro ao ler XLSX: {e}")
        return todos, indexado

    # Normaliza nomes de colunas
    df.columns = [str(c).strip() for c in df.columns]

    for _, row in df.iterrows():
        reg = {k: (str(v).strip() if pd.notna(v) else "") for k, v in row.items()}
        todos.append(reg)

        chave_nfe = limpar_chave(reg.get("Chave Nfe/Cte", ""))
        seq_raw   = reg.get("Sequencia", "0")

        try:
            seq = int(float(seq_raw)) if seq_raw else 0
        except Exception:
            seq = 0

        if len(chave_nfe) < 44 or seq == 0:
            continue

        indexado[(chave_nfe, seq)] = {
            "cfop":          reg.get("Cfop", "").strip(),
            "cod_item":      reg.get("Cod Item", "").strip(),
            "desc_item":     reg.get("Desc Item", "").strip(),
            "ncm":           reg.get("NCM", "").strip(),
            "nro_documento": reg.get("Nro Documento", "").strip(),
            "razao_social":  reg.get("Razao Social", "").strip(),
            "cnpj":          reg.get("CNPJ-CPF", "").strip(),
            "data_entrada":  reg.get("Data Entrada", "").strip(),
            "data_emissao":  reg.get("Data Emissao", "").strip(),
            "uf":            reg.get("UF", "").strip(),
            "vlr_documento": limpar_valor(reg.get("Vlr Documento", "0")),
            "vlr_produto":   limpar_valor(reg.get("Vlr Produto", "0")),
            "cst_icms":      reg.get("CST ICMS", "").strip(),
            "base_icms":     limpar_valor(reg.get("Base Icms", "0")),
            "perc_icms":     limpar_valor(reg.get("Perc ICms", "0")),
            "vlr_icms":      limpar_valor(reg.get("Vlr Icms", "0")),
            "base_icms_st":  limpar_valor(reg.get("Base Icms St", "0")),
            "vlr_icms_st":   limpar_valor(reg.get("Vlr Icms St", "0")),
            "cst_ipi":       reg.get("CST IPI", "").strip(),
            "base_ipi":      limpar_valor(reg.get("Base Ipi", "0")),
            "perc_ipi":      limpar_valor(reg.get("Perc Ipi", "0")),
            "vlr_ipi":       limpar_valor(reg.get("Vlr Ipi", "0")),
            "perc_pis":      limpar_valor(reg.get("Perc Pis", "0")),
            "vlr_pis":       limpar_valor(reg.get("Vlr Pis", "0")),
            "perc_cofins":   limpar_valor(reg.get("Perc Cofins", "0")),
            "vlr_cofins":    limpar_valor(reg.get("Vlr Cofins", "0")),
        }

    return todos, indexado

# ─────────────────────────────────────────────
# PROCESSAMENTO DO XML
# ─────────────────────────────────────────────
NS = "http://www.portalfiscal.inf.br/nfe"

def tag(nome: str) -> str:
    return f"{{{NS}}}{nome}"

def find(elem, *nomes):
    """Navega filhos pelo nome local, ignorando namespace."""
    atual = elem
    for nome in nomes:
        encontrado = None
        for filho in atual:
            local = filho.tag.split("}")[-1] if "}" in filho.tag else filho.tag
            if local == nome:
                encontrado = filho
                break
        if encontrado is None:
            return None
        atual = encontrado
    return atual

def set_text(elem, *caminho, valor: str):
    """Define o texto de um elemento, criando se não existir."""
    alvo = elem
    for nome in caminho:
        filho = find(alvo, nome)
        if filho is None:
            filho = etree.SubElement(alvo, tag(nome))
        alvo = filho
    alvo.text = valor

def get_text(elem, *caminho) -> str:
    alvo = find(elem, *caminho)
    return (alvo.text or "").strip() if alvo is not None else ""

def processar_xml(
    conteudo_xml: bytes,
    nome_arquivo: str,
    dados_indexados: dict,
    cfops_ativas: set,
    campos_marcados: dict,
) -> tuple:
    """
    Processa um XML de NF-e.
    Retorna (xml_modificado_bytes | None, mensagem, status)
    status: 'ok' | 'info' | 'erro'
    """
    try:
        tree = etree.fromstring(conteudo_xml)
    except Exception as e:
        return None, f"XML inválido: {e}", "erro"

    # Localiza infNFe (pode estar dentro de nfeProc ou direto)
    inf_nfe = find(tree, "NFe", "infNFe") or find(tree, "infNFe")
    if inf_nfe is None:
        return None, "infNFe não encontrado", "erro"

    # Extrai chave da NF-e
    chave_xml = inf_nfe.get("Id", "")
    chave_xml = re.sub(r"[^0-9]", "", chave_xml)
    if len(chave_xml) < 44:
        return None, f"Chave inválida no XML: {chave_xml}", "erro"

    # Verifica se há algum item com CFOP ativo no dicionário
    itens_validos = []
    det_elements = []

    for filho in inf_nfe:
        local = filho.tag.split("}")[-1] if "}" in filho.tag else filho.tag
        if local == "det":
            det_elements.append(filho)

    for det in det_elements:
        n_item_str = det.get("nItem", "0")
        try:
            n_item = int(n_item_str)
        except Exception:
            continue

        chave_busca = (chave_xml, n_item)
        dados = dados_indexados.get(chave_busca)
        if dados is None:
            continue

        cfop_xlsx = dados.get("cfop", "")
        if cfop_xlsx in cfops_ativas:
            itens_validos.append((det, n_item, dados))

    if not itens_validos:
        return None, f"nenhum item com CFOP válido — não alterado", "info"

    # Aplica alterações
    modificado = False
    for det, n_item, dados in itens_validos:
        prod = find(det, "prod")
        imposto = find(det, "imposto")
        if prod is None or imposto is None:
            continue

        # ── CFOP ──
        if campos_marcados.get("cfop"):
            cfop_elem = find(prod, "CFOP")
            if cfop_elem is not None:
                cfop_elem.text = dados["cfop"]
                modificado = True

        # ── NCM ──
        if campos_marcados.get("ncm"):
            ncm_elem = find(prod, "NCM")
            if ncm_elem is not None and dados["ncm"]:
                ncm_elem.text = dados["ncm"]
                modificado = True

        # ── ICMS ──
        if campos_marcados.get("icms"):
            icms_pai = find(imposto, "ICMS")
            if icms_pai is not None:
                # Localiza o filho de ICMS (ICMS00, ICMS10, ICMS20, etc.)
                for icms_filho in icms_pai:
                    local_f = icms_filho.tag.split("}")[-1] if "}" in icms_filho.tag else icms_filho.tag
                    if local_f.startswith("ICMS"):
                        cst_elem = find(icms_filho, "CST")
                        vbc_elem = find(icms_filho, "vBC")
                        picms_elem = find(icms_filho, "pICMS")
                        vicms_elem = find(icms_filho, "vICMS")
                        if cst_elem is not None and dados["cst_icms"]:
                            cst_elem.text = dados["cst_icms"].zfill(2)
                            modificado = True
                        if vbc_elem is not None:
                            vbc_elem.text = fmt(dados["base_icms"])
                            modificado = True
                        if picms_elem is not None:
                            picms_elem.text = fmt(dados["perc_icms"])
                            modificado = True
                        if vicms_elem is not None:
                            vicms_elem.text = fmt(dados["vlr_icms"])
                            modificado = True
                        # ICMS-ST
                        vbcst_elem = find(icms_filho, "vBCST")
                        vicmsst_elem = find(icms_filho, "vICMSST")
                        if vbcst_elem is not None:
                            vbcst_elem.text = fmt(dados["base_icms_st"])
                            modificado = True
                        if vicmsst_elem is not None:
                            vicmsst_elem.text = fmt(dados["vlr_icms_st"])
                            modificado = True
                        break

        # ── IPI ──
        if campos_marcados.get("ipi"):
            ipi_elem = find(imposto, "IPI")
            if ipi_elem is not None:
                ipi_trib = find(ipi_elem, "IPITrib")
                if ipi_trib is not None:
                    cst_ipi = find(ipi_trib, "CST")
                    base_ipi = find(ipi_trib, "vBC") or find(ipi_trib, "qUnid")
                    perc_ipi = find(ipi_trib, "pIPI")
                    vlr_ipi  = find(ipi_trib, "vIPI")
                    if cst_ipi is not None and dados["cst_ipi"]:
                        cst_ipi.text = dados["cst_ipi"].zfill(2)
                        modificado = True
                    if vlr_ipi is not None:
                        vlr_ipi.text = fmt(dados["vlr_ipi"])
                        modificado = True

        # ── PIS ──
        if campos_marcados.get("pis"):
            pis_pai = find(imposto, "PIS")
            if pis_pai is not None:
                for pis_filho in pis_pai:
                    vbc_p = find(pis_filho, "vBC")
                    ppis  = find(pis_filho, "pPIS")
                    vpis  = find(pis_filho, "vPIS")
                    if vbc_p is not None:
                        vbc_p.text = fmt(dados["base_icms"])  # base PIS = base produto
                        modificado = True
                    if ppis is not None:
                        ppis.text = fmt(dados["perc_pis"])
                        modificado = True
                    if vpis is not None:
                        vpis.text = fmt(dados["vlr_pis"])
                        modificado = True
                    break

        # ── COFINS ──
        if campos_marcados.get("cofins"):
            cof_pai = find(imposto, "COFINS")
            if cof_pai is not None:
                for cof_filho in cof_pai:
                    vbc_c  = find(cof_filho, "vBC")
                    pcof   = find(cof_filho, "pCOFINS")
                    vcof   = find(cof_filho, "vCOFINS")
                    if vbc_c is not None:
                        vbc_c.text = fmt(dados["base_icms"])
                        modificado = True
                    if pcof is not None:
                        pcof.text = fmt(dados["perc_cofins"])
                        modificado = True
                    if vcof is not None:
                        vcof.text = fmt(dados["vlr_cofins"])
                        modificado = True
                    break

    if not modificado:
        return None, "nenhuma alteração aplicada", "info"

    xml_out = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", pretty_print=False)
    return xml_out, f"alterado com sucesso ({len(itens_validos)} itens)", "ok"

# ─────────────────────────────────────────────
# INTERFACE STREAMLIT
# ─────────────────────────────────────────────

# Upload XLSX
st.subheader("1. Arquivo de entrada (XLSX do Domínio)")
arquivo_xlsx = st.file_uploader("Selecione o arquivo .xlsx", type=["xlsx"])

# Upload XMLs
st.subheader("2. Arquivos XML de NF-e")
arquivos_xml = st.file_uploader(
    "Selecione um ou mais XMLs",
    type=["xml"],
    accept_multiple_files=True,
)

# Seleção de CFOPs
st.subheader("3. CFOPs a processar")
cfops_disponiveis = [
    "1101","1102","1111","1113","1116","1117","1118","1120","1121","1122",
    "1123","1124","1125","1126","1128","1151","1152","1153","1154","1155",
    "1156","1157","1158","1159","1160","1201","1202","1203","1204","1207",
    "1208","1209","1250","1251","1252","1253","1254","1255","1256","1257",
    "1258","1301","1302","1303","1351","1352","1353","1354","1355","1356",
    "1360","1401","1403","1406","1407","1408","1409","1410","1411","1412",
    "1413","1414","1415","1450","1451","1452","1453","1454","1455","1456",
    "1457","1458","1501","1502","1503","1504","1505","1506","1507","1508",
    "1551","1552","1553","1554","1555","1556","1557","1558","1559","1560",
    "1601","1651","1652","1653","1654","1655","1656","1657","1658","1659",
    "1660","1661","1662","1663","1664","1901","1902","1903","1904","1905",
    "1906","1907","1908","1909","1910","1911","1912","1913","1914","1915",
    "1916","1917","1918","1919","1920","1921","1922","1923","1924","1925",
    "1926","1927","1928","1929","1930","1931","1932","1933","1934","1935",
    "2101","2102","2111","2113","2116","2117","2118","2120","2121","2122",
    "2201","2202","2203","2204","2207","2208","2209","2250","2301","2302",
    "2303","2351","2401","2403","2406","2407","2408","2409","2410","2411",
    "2412","2413","2414","2415","2501","2502","2503","2504","2505","2551",
    "2552","2553","2554","2555","2556","2557","2558","2601","2651","2652",
    "2653","2654","2655","2656","2657","2658","2901","2902","2903","2904",
    "2905","2906","2907","2908","2909","2910","2911","2912","2913","2914",
    "5101","5102","5103","5104","5105","5106","5109","5110","5111","5112",
    "5113","5114","5115","5116","5117","5118","5119","5120","5121","5122",
    "5123","5124","5125","5151","5152","5153","5154","5155","5156","5201",
    "5202","5203","5204","5205","5206","5207","5208","5209","5210","5211",
    "5212","5213","5214","5215","5216","5217","5218","5219","5220","5221",
    "5222","5223","5224","5225","5226","5227","5228","5229","5230","5231",
    "5232","5233","5234","5235","5236","5237","5238","5239","5240","5241",
    "5242","5243","5244","5245","5246","5247","5248","5249","5250","5251",
    "5252","5253","5254","5255","5256","5257","5258","5259","5260","5261",
    "5262","5263","5264","5265","5266","5267","5268","5269","5270","5271",
    "5272","5273","5274","5275","5276","5277","5278","5279","5280","5281",
    "5282","5283","5284","5285","5286","5287","5288","5289","5290","5291",
    "5292","5293","5294","5295","5296","5297","5298","5299","5300","5301",
    "5302","5303","5304","5305","5306","5307","5308","5309","5310","5311",
    "5312","5313","5314","5315","5316","5317","5318","5319","5320","5321",
    "5322","5323","5324","5325","5326","5327","5328","5329","5330","5331",
    "5332","5333","5334","5335","5336","5337","5338","5339","5340","5341",
    "5342","5343","5344","5345","5346","5347","5348","5349","5350","5351",
    "5352","5353","5354","5355","5356","5357","5358","5359","5360","5361",
    "5362","5363","5364","5365","5366","5367","5368","5369","5370","5371",
    "5372","5373","5374","5375","5376","5377","5378","5379","5380","5381",
    "5382","5383","5384","5385","5386","5387","5388","5389","5390","5391",
    "5392","5393","5394","5395","5396","5397","5398","5399","5400","5401",
    "5402","5403","5404","5405","5406","5407","5408","5409","5410","5411",
    "5412","5413","5414","5415","5416","5417","5418","5419","5420","5421",
    "5422","5423","5424","5425","5426","5427","5428","5429","5430","5431",
    "5432","5433","5434","5435","5436","5437","5438","5439","5440","5441",
    "5442","5443","5444","5445","5446","5447","5448","5449","5450","5451",
    "5501","5502","5503","5504","5505","5551","5552","5553","5554","5555",
    "5556","5557","5558","5559","5560","5601","5602","5603","5604","5605",
    "5606","5607","5608","5609","5610","5651","5652","5653","5654","5655",
    "5656","5657","5658","5659","5660","5661","5662","5663","5664","5901",
    "5902","5903","5904","5905","5906","5907","5908","5909","5910","5911",
    "5912","5913","5914","5915","5916","5917","5918","5919","5920","5921",
    "5922","5923","5924","5925","5926","5927","5928","5929","5930","5931",
    "5932","5933","5934","5935","5936","5937","5938","5939","5940","5941",
    "5942","5943","5944","5945","5946","5947","5948","5949","5950","5951",
    "5952","5953","5954","5955","5956","5957","5958","5959","5960","5961",
    "5962","5963","5964","5965","5966","5967","5968","5969","5970","5971",
    "5972","5973","5974","5975","5976","5977","5978","5979","5980","5981",
    "5982","5983","5984","5985","5986","5987","5988","5989","5990","5991",
    "5992","5993","5994","5995","5996","5997","5998","5999","6101","6102",
    "6103","6104","6105","6106","6107","6108","6109","6110","6111","6112",
    "6113","6114","6115","6116","6117","6118","6119","6120","6121","6122",
    "6123","6124","6125","6151","6152","6153","6154","6155","6156","6201",
    "6202","6203","6204","6205","6206","6207","6208","6209","6210","6211",
    "6212","6213","6214","6215","6216","6217","6218","6219","6220","6221",
    "6222","6223","6224","6225","6226","6227","6228","6229","6230","6231",
    "6232","6233","6234","6235","6236","6237","6238","6239","6240","6241",
    "6242","6243","6244","6245","6246","6247","6248","6249","6250","6251",
    "6252","6253","6254","6255","6256","6257","6258","6259","6260","6261",
    "6262","6263","6264","6265","6266","6267","6268","6269","6270","6271",
    "6272","6273","6274","6275","6276","6277","6278","6279","6280","6281",
    "6282","6283","6284","6285","6286","6287","6288","6289","6290","6291",
    "6292","6293","6294","6295","6296","6297","6298","6299","6300","6301",
    "6302","6303","6304","6305","6306","6307","6308","6309","6310","6311",
    "6312","6313","6314","6315","6316","6317","6318","6319","6320","6321",
    "6322","6323","6324","6325","6326","6327","6328","6329","6330","6331",
    "6332","6333","6334","6335","6336","6337","6338","6339","6340","6341",
    "6342","6343","6344","6345","6346","6347","6348","6349","6350","6351",
    "6352","6353","6354","6355","6356","6357","6358","6359","6360","6361",
    "6362","6363","6364","6365","6366","6367","6368","6369","6370","6371",
    "6372","6373","6374","6375","6376","6377","6378","6379","6380","6381",
    "6382","6383","6384","6385","6386","6387","6388","6389","6390","6391",
    "6392","6393","6394","6395","6396","6397","6398","6399","6400","6401",
    "6402","6403","6404","6405","6406","6407","6408","6409","6410","6411",
    "6412","6413","6414","6415","6416","6417","6418","6419","6420","6421",
    "6422","6423","6424","6425","6426","6427","6428","6429","6430","6431",
    "6432","6433","6434","6435","6436","6437","6438","6439","6440","6441",
    "6442","6443","6444","6445","6446","6447","6448","6449","6450","6451",
    "6452","6453","6454","6455","6456","6457","6458","6459","6460","6461",
    "6462","6463","6464","6465","6466","6467","6468","6469","6470","6471",
    "6472","6473","6474","6475","6476","6477","6478","6479","6480","6481",
    "6482","6483","6484","6485","6486","6487","6488","6489","6490","6491",
    "6492","6493","6494","6495","6496","6497","6498","6499","6500","6501",
    "6502","6503","6504","6505","6506","6507","6508","6509","6510","6511",
    "6512","6513","6514","6515","6516","6517","6518","6519","6520","6521",
    "6522","6523","6524","6525","6526","6527","6528","6529","6530","6531",
    "6532","6533","6534","6535","6536","6537","6538","6539","6540","6541",
    "6542","6543","6544","6545","6546","6547","6548","6549","6550","6551",
    "6552","6553","6554","6555","6556","6557","6558","6559","6560","6561",
    "6562","6563","6564","6565","6566","6567","6568","6569","6570","6571",
    "6572","6573","6574","6575","6576","6577","6578","6579","6580","6581",
    "6582","6583","6584","6585","6586","6587","6588","6589","6590","6591",
    "6592","6593","6594","6595","6596","6597","6598","6599","6900","6901",
    "6902","6903","6904","6905","6906","6907","6908","6909","6910","6911",
    "6912","6913","6914","6915","6916","6917","6918","6919","6920","6921",
    "6922","6923","6924","6925","6926","6927","6928","6929","6930","6931",
    "6932","6933","6934","6935","6936","6937","6938","6939","6940","6941",
    "6942","6943","6944","6945","6946","6947","6948","6949","6950","6951",
    "6952","6953","6954","6955","6956","6957","6958","6959","6960","6961",
    "6962","6963","6964","6965","6966","6967","6968","6969","6970","6971",
    "6972","6973","6974","6975","6976","6977","6978","6979","6980","6981",
    "6982","6983","6984","6985","6986","6987","6988","6989","6990","6991",
    "6992","6993","6994","6995","6996","6997","6998","6999","7101","7102",
    "7105","7106","7107","7108","7109","7110","7111","7112","7113","7114",
    "7115","7116","7117","7118","7119","7120","7121","7122","7123","7124",
    "7125","7126","7127","7128","7129","7130","7131","7132","7133","7134",
    "7135","7136","7137","7138","7139","7140","7141","7142","7143","7144",
    "7145","7146","7147","7148","7149","7150","7151","7152","7153","7154",
    "7155","7156","7157","7158","7159","7160","7161","7162","7163","7164",
    "7165","7166","7167","7168","7169","7170","7171","7172","7173","7174",
    "7175","7176","7177","7178","7179","7180","7181","7182","7183","7184",
    "7185","7186","7187","7188","7189","7190","7191","7192","7193","7194",
    "7195","7196","7197","7198","7199","7200","7201","7202","7203","7204",
    "7205","7206","7207","7208","7209","7210","7211","7212","7213","7214",
    "7215","7216","7217","7218","7219","7220","7221","7222","7223","7224",
    "7225","7226","7227","7228","7229","7230","7231","7232","7233","7234",
    "7235","7236","7237","7238","7239","7240","7241","7242","7243","7244",
    "7245","7246","7247","7248","7249","7250","7251","7252","7253","7254",
    "7255","7256","7257","7258","7259","7260","7261","7262","7263","7264",
    "7265","7266","7267","7268","7269","7270","7271","7272","7273","7274",
    "7275","7276","7277","7278","7279","7280","7281","7282","7283","7284",
    "7285","7286","7287","7288","7289","7290","7291","7292","7293","7294",
    "7295","7296","7297","7298","7299","7300","7301","7302","7303","7304",
    "7305","7306","7307","7308","7309","7310","7311","7312","7313","7314",
    "7315","7316","7317","7318","7319","7320","7321","7322","7323","7324",
    "7325","7326","7327","7328","7329","7330","7331","7332","7333","7334",
    "7335","7336","7337","7338","7339","7340","7341","7342","7343","7344",
    "7345","7346","7347","7348","7349","7350","7351","7352","7353","7354",
    "7355","7356","7357","7358","7359","7360","7361","7362","7363","7364",
    "7365","7366","7367","7368","7369","7370","7371","7372","7373","7374",
    "7375","7376","7377","7378","7379","7380","7381","7382","7383","7384",
    "7385","7386","7387","7388","7389","7390","7391","7392","7393","7394",
    "7395","7396","7397","7398","7399","7400","7401","7402","7403","7404",
    "7405","7406","7407","7408","7409","7410","7411","7412","7413","7414",
    "7415","7416","7417","7418","7419","7420","7421","7422","7423","7424",
    "7425","7426","7427","7428","7429","7430","7431","7432","7433","7434",
    "7435","7436","7437","7438","7439","7440","7441","7442","7443","7444",
    "7445","7446","7447","7448","7449","7450","7451","7452","7453","7454",
    "7455","7456","7457","7458","7459","7460","7461","7462","7463","7464",
    "7465","7466","7467","7468","7469","7470","7471","7472","7473","7474",
    "7475","7476","7477","7478","7479","7480","7481","7482","7483","7484",
    "7485","7486","7487","7488","7489","7490","7491","7492","7493","7494",
    "7495","7496","7497","7498","7499","7500","7501","7502","7503","7504",
    "7505","7506","7507","7508","7509","7510","7511","7512","7513","7514",
    "7515","7516","7517","7518","7519","7520","7521","7522","7523","7524",
    "7525","7526","7527","7528","7529","7530","7531","7532","7533","7534",
    "7535","7536","7537","7538","7539","7540","7541","7542","7543","7544",
    "7545","7546","7547","7548","7549","7550","7551","7552","7553","7554",
    "7555","7556","7557","7558","7559","7560","7561","7562","7563","7564",
    "7565","7566","7567","7568","7569","7570","7571","7572","7573","7574",
    "7575","7576","7577","7578","7579","7580","7581","7582","7583","7584",
    "7585","7586","7587","7588","7589","7590","7591","7592","7593","7594",
    "7595","7596","7597","7598","7599","7900","7901","7902","7903","7904",
    "7905","7906","7907","7908","7909","7910","7911","7912","7913","7914",
    "7915","7916","7917","7918","7919","7920","7921","7922","7923","7924",
    "7925","7926","7927","7928","7929","7930","7931","7932","7933","7934",
    "7935","7936","7937","7938","7939","7940","7941","7942","7943","7944",
    "7945","7946","7947","7948","7949","7950","7951","7952","7953","7954",
    "7955","7956","7957","7958","7959","7960","7961","7962","7963","7964",
    "7965","7966","7967","7968","7969","7970","7971","7972","7973","7974",
    "7975","7976","7977","7978","7979","7980","7981","7982","7983","7984",
    "7985","7986","7987","7988","7989","7990","7991","7992","7993","7994",
    "7995","7996","7997","7998","7999",
]

# CFOPs padrão pré-selecionadas (devoluções e transferências)
cfops_default = [
    "1201","1202","1203","1204","1207","1208","1209",
    "1401","1403","1406","1407","1408","1409","1410","1411",
    "2201","2202","2203","2204","2207","2208","2209",
    "2401","2403","2406","2407","2408","2409","2410","2411",
]

cfops_selecionadas = st.multiselect(
    "CFOPs que serão processadas:",
    options=cfops_disponiveis,
    default=[c for c in cfops_default if c in cfops_disponiveis],
)

# Seleção de campos
st.subheader("4. Campos a atualizar no XML")
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    cb_cfop   = st.checkbox("CFOP",   value=True)
with col2:
    cb_icms   = st.checkbox("ICMS",   value=True)
with col3:
    cb_ipi    = st.checkbox("IPI",    value=False)
with col4:
    cb_pis    = st.checkbox("PIS",    value=False)
with col5:
    cb_cofins = st.checkbox("COFINS", value=False)
cb_ncm = st.checkbox("NCM", value=False)

campos_marcados = {
    "cfop":   cb_cfop,
    "icms":   cb_icms,
    "ipi":    cb_ipi,
    "pis":    cb_pis,
    "cofins": cb_cofins,
    "ncm":    cb_ncm,
}

# Botão processar
st.subheader("5. Processar")
if st.button("▶️ Processar XMLs", type="primary"):
    if not arquivo_xlsx:
        st.error("Selecione o arquivo XLSX.")
        st.stop()
    if not arquivos_xml:
        st.error("Selecione ao menos um XML.")
        st.stop()
    if not cfops_selecionadas:
        st.error("Selecione ao menos uma CFOP.")
        st.stop()

    cfops_ativas = set(cfops_selecionadas)

    with st.spinner("Lendo XLSX..."):
        _, dados_indexados = ler_xlsx(arquivo_xlsx.read())

    if not dados_indexados:
        st.error("Nenhum item válido encontrado no XLSX. Verifique o arquivo.")
        st.stop()

    st.info(f"XLSX carregado: {len(dados_indexados)} itens indexados.")

    # Debug opcional
    with st.expander("🔍 Primeiros 5 itens indexados (debug)"):
        for i, ((ch, seq), d) in enumerate(list(dados_indexados.items())[:5]):
            st.code(f"Chave: {ch} | Seq: {seq} | CFOP: {d['cfop']} | Item: {d['cod_item']}")

    resultados = []
    xmls_modificados = {}

    progress = st.progress(0)
    total = len(arquivos_xml)

    for idx, arq in enumerate(arquivos_xml):
        conteudo = arq.read()
        xml_out, msg, status = processar_xml(
            conteudo,
            arq.name,
            dados_indexados,
            cfops_ativas,
            campos_marcados,
        )
        resultados.append((arq.name, msg, status))
        if xml_out:
            xmls_modificados[arq.name] = xml_out
        progress.progress((idx + 1) / total)

    progress.empty()

    # Exibe resultados
    st.subheader("Resultados")
    for nome, msg, status in resultados:
        if status == "ok":
            st.success(f"✅ {nome}: {msg}")
        elif status == "info":
            st.info(f"ℹ️ {nome}: {msg}")
        else:
            st.error(f"❌ {nome}: {msg}")

    # Download
    if xmls_modificados:
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
                file_name="xmls_modificados.zip",
                mime="application/zip",
            )
    else:
        st.warning("Nenhum XML foi modificado.")
