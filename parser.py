"""
parser.py
---------
Responsável por receber o arquivo de entrada (XML, JSON ou Excel/XLSX)
e retornar um dicionário padronizado com os dados da DUIMP.
Flexível: se um campo não existir, retorna valor padrão (None ou 0.0).
"""

import json
import re
from pathlib import Path
import xml.etree.ElementTree as ET
import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────────────────────────────────────

def _safe_float(value, default=0.0):
    """Converte para float com segurança; retorna default se falhar."""
    if value is None:
        return default
    try:
        return float(str(value).replace(",", ".").strip())
    except (ValueError, TypeError):
        return default


def _safe_str(value, default=""):
    if value is None:
        return default
    return str(value).strip()


# ─────────────────────────────────────────────────────────────────────────────
# PARSER XML
# ─────────────────────────────────────────────────────────────────────────────

def _parse_xml(filepath: str) -> dict:
    tree = ET.parse(filepath)
    root = tree.getroot()

    # Remove namespaces para facilitar busca
    for elem in root.iter():
        if "}" in elem.tag:
            elem.tag = elem.tag.split("}", 1)[1]

    def find_text(parent, tag, default=""):
        el = parent.find(f".//{tag}")
        return el.text.strip() if el is not None and el.text else default

    # ── Cabeçalho ──
    header = {
        "numero_duimp":    find_text(root, "numeroDuimp") or find_text(root, "NumeroDUIM"),
        "cnpj_importador": find_text(root, "cnpjImportador"),
        "nome_importador": find_text(root, "nomeImportador"),
        "cnpj_adquirente": find_text(root, "cnpjAdquirente"),
        "nome_adquirente": find_text(root, "nomeAdquirente"),
        "ref_capital":     find_text(root, "refCapitalTrade"),
        "ref_cliente":     find_text(root, "refCliente"),
        "fatura":          find_text(root, "fatura"),
        "conhecimento":    find_text(root, "conhecimentoCarga"),
        "container":       find_text(root, "container"),
        "navio":           find_text(root, "navio"),
        "data_embarque":   find_text(root, "dataEmbarque"),
        "data_chegada":    find_text(root, "dataChegada"),
        "valor_fob_brl":   _safe_float(find_text(root, "valorFobBrl")),
        "valor_frete_brl": _safe_float(find_text(root, "valorFreteBrl")),
        "valor_aduaneiro": _safe_float(find_text(root, "valorAduaneiro")),
        "taxa_cambio":     _safe_float(find_text(root, "taxaCambio"), 1.0),
        "canal":           find_text(root, "canal"),
        "pais_procedencia":find_text(root, "paisProcedencia"),
    }

    # ── Adições ──
    adicoes = []
    # Tenta diferentes nomes de tag para as adições
    items = (root.findall(".//adicao") or
             root.findall(".//item") or
             root.findall(".//Item") or
             root.findall(".//Adicao"))

    for idx, item in enumerate(items, start=1):
        def f(tag, default=""):
            return find_text(item, tag, default)

        adicao = {
            "numero_adicao":  _safe_str(f("numeroAdicao") or f("NumeroAdicao"), str(idx).zfill(5)),
            "ncm":            _safe_str(f("ncm") or f("NCM")),
            "descricao":      _safe_str(f("descricao") or f("descricaoProduto")),
            "detalhamento":   _safe_str(f("detalhamento") or f("informacaoComplementar")),
            "quantidade":     _safe_float(f("quantidade")),
            "unidade":        _safe_str(f("unidade"), "UN"),
            "valor_unitario": _safe_float(f("valorUnitario")),
            "valor_total_usd":_safe_float(f("valorTotalUsd") or f("valorCondicaoVenda")),
            "valor_aduaneiro":_safe_float(f("valorAduaneiro")),
            "peso_liquido":   _safe_float(f("pesoLiquido")),
            "ii_aliquota":    _safe_float(f("iiAliquota")),
            "ii_valor":       _safe_float(f("iiValor")),
            "ipi_aliquota":   _safe_float(f("ipiAliquota")),
            "ipi_valor":      _safe_float(f("ipiValor")),
            "pis_aliquota":   _safe_float(f("pisAliquota")),
            "pis_valor":      _safe_float(f("pisValor")),
            "cofins_aliquota":_safe_float(f("cofinsAliquota")),
            "cofins_valor":   _safe_float(f("cofinsValor")),
            "icms_aliquota":  _safe_float(f("icmsAliquota")),
            "icms_valor":     _safe_float(f("icmsValor")),
            "regime_ii":      _safe_str(f("regimeTributacaoII")),
            "regime_ipi":     _safe_str(f("regimeTributacaoIPI")),
            "fabricante":     _safe_str(f("fabricante") or f("nomeFabricante")),
            "pais_origem":    _safe_str(f("paisOrigem")),
            "aplicacao":      _safe_str(f("aplicacao"), "Revenda"),
            "condicao":       _safe_str(f("condicao"), "Nova"),
        }
        adicoes.append(adicao)

    return {"header": header, "adicoes": adicoes}


# ─────────────────────────────────────────────────────────────────────────────
# PARSER JSON
# ─────────────────────────────────────────────────────────────────────────────

def _parse_json(filepath: str) -> dict:
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Suporta estrutura com chave "duimp" ou diretamente o objeto
    root = data.get("duimp", data)

    header_raw = root.get("cabecalho", root.get("header", root))
    header = {
        "numero_duimp":    _safe_str(header_raw.get("numeroDuimp", header_raw.get("numero", ""))),
        "cnpj_importador": _safe_str(header_raw.get("cnpjImportador", "")),
        "nome_importador": _safe_str(header_raw.get("nomeImportador", "")),
        "cnpj_adquirente": _safe_str(header_raw.get("cnpjAdquirente", "")),
        "nome_adquirente": _safe_str(header_raw.get("nomeAdquirente", "")),
        "ref_capital":     _safe_str(header_raw.get("refCapitalTrade", "")),
        "ref_cliente":     _safe_str(header_raw.get("refCliente", "")),
        "fatura":          _safe_str(header_raw.get("fatura", "")),
        "conhecimento":    _safe_str(header_raw.get("conhecimentoCarga", "")),
        "container":       _safe_str(header_raw.get("container", "")),
        "navio":           _safe_str(header_raw.get("navio", "")),
        "data_embarque":   _safe_str(header_raw.get("dataEmbarque", "")),
        "data_chegada":    _safe_str(header_raw.get("dataChegada", "")),
        "valor_fob_brl":   _safe_float(header_raw.get("valorFobBrl", 0)),
        "valor_frete_brl": _safe_float(header_raw.get("valorFreteBrl", 0)),
        "valor_aduaneiro": _safe_float(header_raw.get("valorAduaneiro", 0)),
        "taxa_cambio":     _safe_float(header_raw.get("taxaCambio", 1.0)),
        "canal":           _safe_str(header_raw.get("canal", "")),
        "pais_procedencia":_safe_str(header_raw.get("paisProcedencia", "")),
    }

    adicoes_raw = root.get("adicoes", root.get("itens", root.get("items", [])))
    adicoes = []
    for idx, item in enumerate(adicoes_raw, start=1):
        adicao = {
            "numero_adicao":  _safe_str(item.get("numeroAdicao", str(idx).zfill(5))),
            "ncm":            _safe_str(item.get("ncm", "")),
            "descricao":      _safe_str(item.get("descricao", item.get("descricaoProduto", ""))),
            "detalhamento":   _safe_str(item.get("detalhamento", item.get("informacaoComplementar", ""))),
            "quantidade":     _safe_float(item.get("quantidade", 0)),
            "unidade":        _safe_str(item.get("unidade", "UN")),
            "valor_unitario": _safe_float(item.get("valorUnitario", 0)),
            "valor_total_usd":_safe_float(item.get("valorTotalUsd", item.get("valorCondicaoVenda", 0))),
            "valor_aduaneiro":_safe_float(item.get("valorAduaneiro", 0)),
            "peso_liquido":   _safe_float(item.get("pesoLiquido", 0)),
            "ii_aliquota":    _safe_float(item.get("iiAliquota", 0)),
            "ii_valor":       _safe_float(item.get("iiValor", 0)),
            "ipi_aliquota":   _safe_float(item.get("ipiAliquota", 0)),
            "ipi_valor":      _safe_float(item.get("ipiValor", 0)),
            "pis_aliquota":   _safe_float(item.get("pisAliquota", 0)),
            "pis_valor":      _safe_float(item.get("pisValor", 0)),
            "cofins_aliquota":_safe_float(item.get("cofinsAliquota", 0)),
            "cofins_valor":   _safe_float(item.get("cofinsValor", 0)),
            "icms_aliquota":  _safe_float(item.get("icmsAliquota", 0)),
            "icms_valor":     _safe_float(item.get("icmsValor", 0)),
            "regime_ii":      _safe_str(item.get("regimeII", "RECOLHIMENTO INTEGRAL")),
            "regime_ipi":     _safe_str(item.get("regimeIPI", "RECOLHIMENTO INTEGRAL")),
            "fabricante":     _safe_str(item.get("fabricante", "")),
            "pais_origem":    _safe_str(item.get("paisOrigem", "China, República Popular - CN")),
            "aplicacao":      _safe_str(item.get("aplicacao", "Revenda")),
            "condicao":       _safe_str(item.get("condicao", "Nova")),
        }
        adicoes.append(adicao)

    return {"header": header, "adicoes": adicoes}


# ─────────────────────────────────────────────────────────────────────────────
# PARSER EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def _parse_excel(filepath: str) -> dict:
    """
    Espera uma planilha Excel com:
    - Aba "Cabecalho" (ou primeira aba com par chave/valor)
    - Aba "Adicoes" (ou aba com tabela de adições)
    Se não encontrar as abas, tenta interpretar a primeira aba como tabela de adições.
    """
    xl = pd.ExcelFile(filepath)
    sheet_names_lower = [s.lower() for s in xl.sheet_names]

    # ── Cabeçalho ──
    header = {
        "numero_duimp": "", "cnpj_importador": "", "nome_importador": "",
        "cnpj_adquirente": "", "nome_adquirente": "", "ref_capital": "",
        "ref_cliente": "", "fatura": "", "conhecimento": "", "container": "",
        "navio": "", "data_embarque": "", "data_chegada": "", "valor_fob_brl": 0.0,
        "valor_frete_brl": 0.0, "valor_aduaneiro": 0.0, "taxa_cambio": 1.0,
        "canal": "", "pais_procedencia": "",
    }

    cab_sheet = None
    for name in ["cabecalho", "cabeçalho", "header", "dados", "resumo"]:
        if name in sheet_names_lower:
            cab_sheet = xl.sheet_names[sheet_names_lower.index(name)]
            break

    if cab_sheet:
        df_cab = pd.read_excel(filepath, sheet_name=cab_sheet, header=None)
        # Lê par (chave, valor) de 2 colunas
        mapping = {
            "numero": "numero_duimp", "duimp": "numero_duimp",
            "cnpj importador": "cnpj_importador",
            "nome importador": "nome_importador",
            "cnpj adquirente": "cnpj_adquirente",
            "nome adquirente": "nome_adquirente",
            "ref capital": "ref_capital", "ref. capital": "ref_capital",
            "ref cliente": "ref_cliente", "ref. cliente": "ref_cliente",
            "fatura": "fatura", "conhecimento": "conhecimento",
            "container": "container", "navio": "navio",
            "data embarque": "data_embarque", "data chegada": "data_chegada",
            "valor fob": "valor_fob_brl", "valor frete": "valor_frete_brl",
            "valor aduaneiro": "valor_aduaneiro", "taxa cambio": "taxa_cambio",
            "canal": "canal", "pais procedencia": "pais_procedencia",
        }
        for _, row in df_cab.iterrows():
            if len(row) >= 2 and pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]):
                key_raw = str(row.iloc[0]).lower().strip()
                for k, v in mapping.items():
                    if k in key_raw:
                        val = row.iloc[1]
                        if v in ("valor_fob_brl", "valor_frete_brl", "valor_aduaneiro", "taxa_cambio"):
                            header[v] = _safe_float(val)
                        else:
                            header[v] = _safe_str(val)
                        break

    # ── Adições ──
    add_sheet = None
    for name in ["adicoes", "adições", "itens", "items", "mercadorias"]:
        if name in sheet_names_lower:
            add_sheet = xl.sheet_names[sheet_names_lower.index(name)]
            break

    if add_sheet is None:
        # Usa primeira aba que não seja cabecalho
        for sn in xl.sheet_names:
            if sn.lower() not in ["cabecalho", "cabeçalho", "header"]:
                add_sheet = sn
                break
        if add_sheet is None:
            add_sheet = xl.sheet_names[0]

    df = pd.read_excel(filepath, sheet_name=add_sheet)
    df.columns = [str(c).lower().strip().replace(" ", "_").replace("ã", "a")
                  .replace("é", "e").replace("ç", "c").replace("ó", "o")
                  .replace("ú", "u").replace("á", "a").replace("í", "i")
                  for c in df.columns]

    col_map = {
        "adicao": "numero_adicao", "numero_adicao": "numero_adicao",
        "ncm": "ncm",
        "descricao": "descricao", "descricao_produto": "descricao",
        "detalhamento": "detalhamento",
        "qtde": "quantidade", "quantidade": "quantidade",
        "unidade": "unidade",
        "valor_unitario": "valor_unitario", "vl_unitario": "valor_unitario",
        "valor_total_usd": "valor_total_usd", "valor_total": "valor_total_usd",
        "valor_aduaneiro": "valor_aduaneiro",
        "peso_liquido": "peso_liquido",
        "ii_%": "ii_aliquota", "ii_aliq": "ii_aliquota", "ii_aliquota": "ii_aliquota",
        "ii_r$": "ii_valor", "ii_valor": "ii_valor",
        "ipi_%": "ipi_aliquota", "ipi_aliq": "ipi_aliquota", "ipi_aliquota": "ipi_aliquota",
        "ipi_r$": "ipi_valor", "ipi_valor": "ipi_valor",
        "pis_%": "pis_aliquota", "pis_aliq": "pis_aliquota", "pis_aliquota": "pis_aliquota",
        "pis_r$": "pis_valor", "pis_valor": "pis_valor",
        "cofins_%": "cofins_aliquota", "cofins_aliq": "cofins_aliquota", "cofins_aliquota": "cofins_aliquota",
        "cofins_r$": "cofins_valor", "cofins_valor": "cofins_valor",
        "icms_%": "icms_aliquota", "icms_aliq": "icms_aliquota", "icms_aliquota": "icms_aliquota",
        "icms_r$": "icms_valor", "icms_valor": "icms_valor",
        "fabricante": "fabricante", "pais_origem": "pais_origem",
        "aplicacao": "aplicacao", "condicao": "condicao",
    }

    adicoes = []
    for idx, row in df.iterrows():
        adicao = {
            "numero_adicao": "", "ncm": "", "descricao": "", "detalhamento": "",
            "quantidade": 0.0, "unidade": "UN", "valor_unitario": 0.0,
            "valor_total_usd": 0.0, "valor_aduaneiro": 0.0, "peso_liquido": 0.0,
            "ii_aliquota": 0.0, "ii_valor": 0.0,
            "ipi_aliquota": 0.0, "ipi_valor": 0.0,
            "pis_aliquota": 0.0, "pis_valor": 0.0,
            "cofins_aliquota": 0.0, "cofins_valor": 0.0,
            "icms_aliquota": 0.0, "icms_valor": 0.0,
            "regime_ii": "RECOLHIMENTO INTEGRAL",
            "regime_ipi": "RECOLHIMENTO INTEGRAL",
            "fabricante": "", "pais_origem": "", "aplicacao": "Revenda", "condicao": "Nova",
        }
        for col_raw, col_norm in col_map.items():
            if col_raw in df.columns:
                val = row.get(col_raw)
                if col_norm in ("numero_adicao", "ncm", "descricao", "detalhamento",
                                "unidade", "regime_ii", "regime_ipi", "fabricante",
                                "pais_origem", "aplicacao", "condicao"):
                    adicao[col_norm] = _safe_str(val)
                else:
                    adicao[col_norm] = _safe_float(val)

        if not adicao["numero_adicao"]:
            adicao["numero_adicao"] = str(idx + 1).zfill(5)

        # Só adiciona linha se tiver NCM ou descrição
        if adicao["ncm"] or adicao["descricao"]:
            adicoes.append(adicao)

    return {"header": header, "adicoes": adicoes}


# ─────────────────────────────────────────────────────────────────────────────
# ENTRADA PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def parse_file(filepath: str) -> dict:
    """
    Detecta o tipo do arquivo e chama o parser adequado.
    Retorna dict com 'header' e 'adicoes'.
    Lança ValueError se o formato não for suportado ou arquivo inválido.
    """
    path = Path(filepath)
    ext = path.suffix.lower()

    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {filepath}")

    try:
        if ext == ".xml":
            return _parse_xml(filepath)
        elif ext == ".json":
            return _parse_json(filepath)
        elif ext in (".xlsx", ".xls", ".xlsm"):
            return _parse_excel(filepath)
        else:
            raise ValueError(f"Formato não suportado: '{ext}'. Use XML, JSON ou Excel.")
    except (ET.ParseError, json.JSONDecodeError) as e:
        raise ValueError(f"Arquivo inválido ou corrompido: {e}") from e
    except Exception as e:
        if "não suportado" in str(e) or "não encontrado" in str(e):
            raise
        raise ValueError(f"Erro ao processar arquivo: {e}") from e
