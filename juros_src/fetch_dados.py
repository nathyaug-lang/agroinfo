"""Conjunto de fetchers de séries macro, todas com cache local em JSON.

Cada função retorna uma lista de tuplas `(data_iso, valor)` ordenada do
mais antigo ao mais recente. As caches ficam em `.cache/<nome>.json` na
pasta do projeto, com freshness configurável.

Fontes públicas usadas:
- BCB SGS  (Sistema Gerenciador de Séries Temporais)
- BCB Olinda Expectativas (Boletim Focus)
- Yahoo Finance v8 (ICE U.S. Dollar Index — DX-Y.NYB)
- World Bank Open Data (balança comercial USA — frequência anual)

Em caso de falha de rede usa o cache mesmo que velho; se o cache também
não existir, propaga a exceção (o build.py decide o que fazer).
"""

from __future__ import annotations

import datetime as dt
import json
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CACHE_DIR = ROOT / ".cache"
CACHE_DIR.mkdir(exist_ok=True)

UA = "Mozilla/5.0 (Innovagro Market Intelligence)"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _http_json(url: str, timeout: int = 30) -> object:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def _cache_path(nome: str) -> Path:
    return CACHE_DIR / f"{nome}.json"


def _cache_age_h(nome: str) -> float | None:
    p = _cache_path(nome)
    if not p.exists():
        return None
    return (time.time() - p.stat().st_mtime) / 3600


def _save_cache(nome: str, dados: list[tuple[str, float]]) -> None:
    _cache_path(nome).write_text(
        json.dumps(dados, ensure_ascii=False), encoding="utf-8"
    )


def _load_cache(nome: str) -> list[tuple[str, float]] | None:
    p = _cache_path(nome)
    if not p.exists():
        return None
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
        return [(d, float(v)) for d, v in raw]
    except Exception:
        return None


def _fetch_or_cache(nome: str, fetcher, max_h: float) -> list[tuple[str, float]]:
    """Padrão: usa cache fresco; senão tenta fetcher; em falha usa cache velho."""
    age = _cache_age_h(nome)
    if age is not None and age < max_h:
        cached = _load_cache(nome)
        if cached:
            return cached
    try:
        dados = fetcher()
        if dados:
            _save_cache(nome, dados)
        return dados
    except Exception as e:
        cached = _load_cache(nome)
        if cached:
            print(f"[fetch_dados] {nome}: API indisponível ({e}). Usando cache antigo.")
            return cached
        raise


# ---------------------------------------------------------------------------
# BCB SGS — backbone de quase todas as séries do Brasil
# ---------------------------------------------------------------------------

def _bcb_sgs(codigo: int, ini: str, fim: str, timeout: int = 30) -> list[dict]:
    url = (
        f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados"
        f"?formato=json&dataInicial={ini}&dataFinal={fim}"
    )
    last: Exception | None = None
    for tent in range(3):
        try:
            return _http_json(url, timeout=timeout)
        except (urllib.error.URLError, TimeoutError) as e:
            last = e
            time.sleep(2 ** tent)
    raise RuntimeError(f"BCB SGS {codigo}: {last}")


def _parse_br_data(s: str) -> str:
    d, m, y = s.split("/")
    return f"{y}-{m}-{d}"


def _sgs_serie(codigo: int, ini_iso: str = "2003-01-01") -> list[tuple[str, float]]:
    """Coleta a série inteira do BCB em chunks de 5 anos."""
    hoje = dt.date.today()
    cur = dt.date.fromisoformat(ini_iso)
    out: list[dict] = []
    while cur <= hoje:
        nxt = dt.date(min(cur.year + 5, hoje.year + 1), 1, 1) - dt.timedelta(days=1)
        if nxt > hoje:
            nxt = hoje
        out.extend(_bcb_sgs(codigo, cur.strftime("%d/%m/%Y"), nxt.strftime("%d/%m/%Y")))
        cur = nxt + dt.timedelta(days=1)
    return [(_parse_br_data(d["data"]), float(d["valor"])) for d in out]


# Quatro séries do BCB com fácil acesso ------------------------------------

def divida_pib_br(max_h: float = 24) -> list[tuple[str, float]]:
    """DBGG % PIB (BCB SGS 13762). Mensal."""
    return _fetch_or_cache("divida_pib_br", lambda: _sgs_serie(13762), max_h)


def usd_brl(max_h: float = 12) -> list[tuple[str, float]]:
    """USD/BRL PTAX venda (BCB SGS 1). Diária — vamos amostrar mensal no build."""
    return _fetch_or_cache("usd_brl", lambda: _sgs_serie(1), max_h)


def balanca_comercial_br(max_h: float = 24) -> list[tuple[str, float]]:
    """Balança Comercial Brasil — saldo mensal em US$ milhões (SGS 22707)."""
    return _fetch_or_cache("balanca_br", lambda: _sgs_serie(22707), max_h)


def fluxo_capital_estrangeiro(max_h: float = 24) -> list[tuple[str, float]]:
    """Investimentos estrangeiros em portfolio — passivos líquidos
    (SGS 11759). Mensal, US$ milhões."""
    return _fetch_or_cache(
        "fluxo_capital_br", lambda: _sgs_serie(11759), max_h
    )


# ---------------------------------------------------------------------------
# BCB Olinda — Expectativas Focus
# ---------------------------------------------------------------------------

def focus_expectativas(max_h: float = 12) -> dict:
    """Snapshot atual do Focus para os principais indicadores.

    Devolve dict no formato:
        {"Selic": {"2026": {"media":..., "mediana":..., ...}, "2027": {...}}, ...}
    Captura a leitura mais recente (1 dia) para cada (Indicador, DataReferencia).
    """
    nome = "focus_expectativas"
    age = _cache_age_h(nome)
    if age is not None and age < max_h:
        try:
            return json.loads(_cache_path(nome).read_text(encoding="utf-8"))
        except Exception:
            pass

    indicadores = ["Selic", "IPCA", "PIB Total", "Câmbio"]
    out: dict[str, dict] = {}
    try:
        for ind in indicadores:
            ind_url = urllib.parse.quote(ind, safe="")  # type: ignore  # noqa
            # Pega últimas 200 obs do indicador, ordenadas decrescente
            # Olinda usa OData
            url = (
                "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/"
                "odata/ExpectativasMercadoAnuais"
                f"?%24top=200"
                f"&%24filter=Indicador%20eq%20'{ind_url}'"
                f"&%24orderby=Data%20desc"
                f"&%24format=json"
            )
            j = _http_json(url, timeout=30)
            por_ano: dict[str, dict] = {}
            for r in j.get("value", []):
                ref = str(r.get("DataReferencia"))
                # mantém só a primeira (mais recente) leitura por ano de referência
                if ref not in por_ano:
                    por_ano[ref] = {
                        "data": r.get("Data"),
                        "media": r.get("Media"),
                        "mediana": r.get("Mediana"),
                        "minimo": r.get("Minimo"),
                        "maximo": r.get("Maximo"),
                        "respondentes": r.get("numeroRespondentes"),
                    }
            out[ind] = por_ano
        _cache_path(nome).write_text(
            json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return out
    except Exception as e:
        if _cache_path(nome).exists():
            print(f"[fetch_dados] Focus indisponível ({e}). Usando cache antigo.")
            return json.loads(_cache_path(nome).read_text(encoding="utf-8"))
        raise


import urllib.parse  # noqa: E402  — placed late on purpose, after function definition


# ---------------------------------------------------------------------------
# Yahoo Finance — DXY (ICE U.S. Dollar Index)
# ---------------------------------------------------------------------------

def dollar_index(max_h: float = 24) -> list[tuple[str, float]]:
    """Dólar Index (DXY, ICE) mensal via Yahoo Finance v8."""
    def fetcher() -> list[tuple[str, float]]:
        p1 = int(dt.datetime(2003, 1, 1).timestamp())
        p2 = int(dt.datetime.now().timestamp())
        url = (
            "https://query1.finance.yahoo.com/v8/finance/chart/DX-Y.NYB"
            f"?period1={p1}&period2={p2}&interval=1mo"
        )
        j = _http_json(url, timeout=30)
        res = j["chart"]["result"][0]
        ts = res["timestamp"]
        close = res["indicators"]["quote"][0]["close"]
        out = []
        for t, c in zip(ts, close):
            if c is None:
                continue
            d_iso = dt.date.fromtimestamp(t).isoformat()
            out.append((d_iso, float(c)))
        return out
    return _fetch_or_cache("dollar_index", fetcher, max_h)


# ---------------------------------------------------------------------------
# World Bank — balança comercial EUA (anual)
# ---------------------------------------------------------------------------

def balanca_comercial_us(max_h: float = 24 * 7) -> list[tuple[str, float]]:
    """Net trade in goods and services — USA, US$ correntes (anual).

    World Bank indicator NE.RSB.GNFS.CD. Saldo positivo = superávit.
    """
    def fetcher() -> list[tuple[str, float]]:
        url = (
            "https://api.worldbank.org/v2/country/USA/indicator/NE.RSB.GNFS.CD"
            "?format=json&per_page=60&date=2003:2026"
        )
        j = _http_json(url, timeout=30)
        if not isinstance(j, list) or len(j) < 2:
            return []
        rows = j[1] or []
        out = []
        for r in rows:
            yr = r.get("date")
            v = r.get("value")
            if yr and v is not None:
                out.append((f"{yr}-12-31", float(v)))
        out.sort()
        return out
    return _fetch_or_cache("balanca_us", fetcher, max_h)


# ---------------------------------------------------------------------------
# CLI rápido — só para debugar
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("Testando fontes...\n")
    for nome, fn in (
        ("Dívida/PIB BR (SGS 13762)", divida_pib_br),
        ("USD/BRL PTAX (SGS 1)", usd_brl),
        ("Balança BR (SGS 22707)", balanca_comercial_br),
        ("Fluxo capital (SGS 11759)", fluxo_capital_estrangeiro),
        ("Dollar Index DXY (Yahoo)", dollar_index),
        ("Balança US (World Bank)", balanca_comercial_us),
    ):
        try:
            d = fn()
            print(f"[OK] {nome}: {len(d)} pontos, [0]={d[0]}, [-1]={d[-1]}")
        except Exception as e:
            print(f"[FAIL] {nome}: {e}")
    try:
        f = focus_expectativas()
        print(f"[OK] Focus: {len(f)} indicadores, ex Selic: {list(f.get('Selic', {}).items())[:2]}")
    except Exception as e:
        print(f"[FAIL] Focus: {e}")
