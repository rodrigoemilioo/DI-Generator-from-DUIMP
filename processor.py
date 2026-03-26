"""
processor.py
------------
Recebe o dict bruto do parser e aplica a lógica de negócio:
- Cálculo automático de impostos quando não informados
- Enriquecimento das adições com alíquotas padrão por NCM
- Cálculo do valor aduaneiro por adição
- Geração de totais consolidados
- Validação e normalização de campos
"""

from __future__ import annotations

# Alíquotas padrão de II por prefixo de NCM (6 dígitos)
# Fonte: TEC – Tarifa Externa Comum (valores aproximados para autopeças)
_II_ALIQUOTAS_PADRAO: dict[str, float] = {
    "851220": 18.0,   # Faróis e lanternas
    "851220": 18.0,
    "870829": 14.0,   # Peças diversas para veículos
    "870821": 14.0,
    "870830": 14.0,
    "401110": 16.0,   # Pneumáticos
    "840991": 14.0,
    "150920": 0.0,    # Azeite de oliva extra virgem
    "150910": 0.0,
}

# Alíquotas de IPI por NCM (capítulo)
_IPI_ALIQUOTAS_PADRAO: dict[str, float] = {
    "8512": 5.0,      # Equipamentos elétricos para veículos
    "8708": 5.0,      # Peças e acessórios para veículos
    "4011": 10.0,     # Pneumáticos
    "1509": 0.0,      # Azeite
}

# PIS/COFINS padrão importação (autopeças – art. 8 Lei 10.865)
_PIS_PADRAO  = 2.10
_COFINS_PADRAO = 9.65

# ICMS padrão SC importação (18%)
_ICMS_PADRAO = 18.0


def _get_ii_aliquota(ncm: str) -> float:
    ncm_clean = ncm.replace(".", "").replace(" ", "")
    for prefix, aliq in _II_ALIQUOTAS_PADRAO.items():
        if ncm_clean.startswith(prefix):
            return aliq
    return 18.0  # fallback


def _get_ipi_aliquota(ncm: str) -> float:
    ncm_clean = ncm.replace(".", "").replace(" ", "")
    for prefix, aliq in _IPI_ALIQUOTAS_PADRAO.items():
        if ncm_clean.startswith(prefix):
            return aliq
    return 5.0  # fallback


def _calcular_ii(valor_aduaneiro: float, aliquota: float) -> float:
    return round(valor_aduaneiro * aliquota / 100, 2)


def _calcular_ipi(valor_aduaneiro: float, ii: float, aliquota: float) -> float:
    """IPI = (VA + II) * aliq%"""
    return round((valor_aduaneiro + ii) * aliquota / 100, 2)


def _calcular_pis_cofins(valor_aduaneiro: float, ii: float, ipi: float,
                          aliq_pis: float, aliq_cofins: float) -> tuple[float, float]:
    """
    Base PIS/COFINS = VA + II + IPI (simplificado).
    Fórmula real inclui AFRMM, etc., mas esta cobre o caso geral.
    """
    base = valor_aduaneiro + ii + ipi
    pis    = round(base * aliq_pis / 100, 2)
    cofins = round(base * aliq_cofins / 100, 2)
    return pis, cofins


def _calcular_icms(valor_aduaneiro: float, ii: float, ipi: float,
                   pis: float, cofins: float, aliquota: float) -> float:
    """
    ICMS importação SC:
    Base ICMS = (VA + II + IPI + PIS + COFINS + despesas aduaneiras) / (1 - aliq%)
    Simplificado sem despesas aduaneiras.
    """
    soma = valor_aduaneiro + ii + ipi + pis + cofins
    base = round(soma / (1 - aliquota / 100), 2)
    return round(base * aliquota / 100, 2)


# ─────────────────────────────────────────────────────────────────────────────
# PROCESSADOR PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def process_data(raw: dict, taxa_cambio_override: float | None = None) -> dict:
    """
    Recebe dict {'header': {...}, 'adicoes': [...]} do parser.
    Retorna dict enriquecido com impostos calculados e totais.

    Parâmetro opcional taxa_cambio_override: se informado, usa essa taxa
    para converter USD → BRL nas adições sem valor aduaneiro em BRL.
    """
    header = raw.get("header", {})
    adicoes_raw = raw.get("adicoes", [])

    # Taxa de câmbio: do header ou override
    taxa_cambio = taxa_cambio_override or header.get("taxa_cambio", 1.0) or 1.0

    adicoes_processadas = []

    totais = {
        "valor_aduaneiro": 0.0,
        "ii": 0.0,
        "ipi": 0.0,
        "pis": 0.0,
        "cofins": 0.0,
        "icms": 0.0,
        "total_tributos": 0.0,
        "quantidade_total": 0.0,
        "peso_liquido_total": 0.0,
    }

    for idx, ad in enumerate(adicoes_raw, start=1):
        item = dict(ad)  # cópia para não mutar o original

        # Número da adição formatado
        num = item.get("numero_adicao", "")
        if not num or num == "0":
            item["numero_adicao"] = str(idx).zfill(5)

        # ── Valor aduaneiro por adição ──
        va = item.get("valor_aduaneiro", 0.0)
        if va == 0.0:
            # Calcula a partir do valor total em USD * taxa de câmbio
            val_usd = item.get("valor_total_usd", 0.0)
            va = round(val_usd * taxa_cambio, 2)
            item["valor_aduaneiro"] = va

        ncm = item.get("ncm", "")

        # ── II ──
        ii_aliq = item.get("ii_aliquota", 0.0)
        ii_val  = item.get("ii_valor", 0.0)
        if ii_aliq == 0.0 and ii_val == 0.0:
            # Verifica regime
            regime_ii = item.get("regime_ii", "RECOLHIMENTO INTEGRAL").upper()
            if "SUSPENSO" in regime_ii or "ISENTO" in regime_ii or "ZERO" in regime_ii:
                ii_aliq = 0.0
                ii_val  = 0.0
            else:
                ii_aliq = _get_ii_aliquota(ncm)
                ii_val  = _calcular_ii(va, ii_aliq)
        elif ii_val == 0.0 and ii_aliq > 0:
            ii_val = _calcular_ii(va, ii_aliq)
        elif ii_aliq == 0.0 and ii_val > 0 and va > 0:
            ii_aliq = round(ii_val / va * 100, 2)
        item["ii_aliquota"] = ii_aliq
        item["ii_valor"]    = ii_val

        # ── IPI ──
        ipi_aliq = item.get("ipi_aliquota", 0.0)
        ipi_val  = item.get("ipi_valor", 0.0)
        if ipi_aliq == 0.0 and ipi_val == 0.0:
            regime_ipi = item.get("regime_ipi", "RECOLHIMENTO INTEGRAL").upper()
            if "SUSPENSO" in regime_ipi or "ISENTO" in regime_ipi or "ZERO" in regime_ipi:
                ipi_aliq = 0.0
                ipi_val  = 0.0
            elif ncm:
                ipi_aliq = _get_ipi_aliquota(ncm)
                ipi_val  = _calcular_ipi(va, ii_val, ipi_aliq)
        elif ipi_val == 0.0 and ipi_aliq > 0:
            ipi_val = _calcular_ipi(va, ii_val, ipi_aliq)
        elif ipi_aliq == 0.0 and ipi_val > 0 and (va + ii_val) > 0:
            ipi_aliq = round(ipi_val / (va + ii_val) * 100, 2)
        item["ipi_aliquota"] = ipi_aliq
        item["ipi_valor"]    = ipi_val

        # ── PIS ──
        pis_aliq = item.get("pis_aliquota", 0.0)
        pis_val  = item.get("pis_valor", 0.0)
        if pis_aliq == 0.0:
            pis_aliq = _PIS_PADRAO
        if pis_val == 0.0:
            pis_val, _ = _calcular_pis_cofins(va, ii_val, ipi_val, pis_aliq, 0)
        item["pis_aliquota"] = pis_aliq
        item["pis_valor"]    = pis_val

        # ── COFINS ──
        cofins_aliq = item.get("cofins_aliquota", 0.0)
        cofins_val  = item.get("cofins_valor", 0.0)
        if cofins_aliq == 0.0:
            cofins_aliq = _COFINS_PADRAO
        if cofins_val == 0.0:
            _, cofins_val = _calcular_pis_cofins(va, ii_val, ipi_val, 0, cofins_aliq)
        item["cofins_aliquota"] = cofins_aliq
        item["cofins_valor"]    = cofins_val

        # ── ICMS ──
        icms_aliq = item.get("icms_aliquota", 0.0)
        icms_val  = item.get("icms_valor", 0.0)
        if icms_aliq == 0.0:
            icms_aliq = _ICMS_PADRAO
        if icms_val == 0.0:
            icms_val = _calcular_icms(va, ii_val, ipi_val, pis_val, cofins_val, icms_aliq)
        item["icms_aliquota"] = icms_aliq
        item["icms_valor"]    = icms_val

        # ── Total por adição ──
        item["total_tributos"] = round(
            ii_val + ipi_val + pis_val + cofins_val + icms_val, 2
        )

        # ── Acumula totais ──
        totais["valor_aduaneiro"]  += va
        totais["ii"]               += ii_val
        totais["ipi"]              += ipi_val
        totais["pis"]              += pis_val
        totais["cofins"]           += cofins_val
        totais["icms"]             += icms_val
        totais["total_tributos"]   += item["total_tributos"]
        totais["quantidade_total"] += item.get("quantidade", 0.0)
        totais["peso_liquido_total"] += item.get("peso_liquido", 0.0)

        adicoes_processadas.append(item)

    # Arredonda totais
    for k in totais:
        totais[k] = round(totais[k], 2)

    return {
        "header": header,
        "adicoes": adicoes_processadas,
        "totais": totais,
        "taxa_cambio_usada": taxa_cambio,
    }
