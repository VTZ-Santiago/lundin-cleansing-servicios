from __future__ import annotations

from datetime import datetime, timezone
import json
from pathlib import Path
from typing import Iterable
import urllib.request


SOURCE_URL = "https://open.er-api.com/v6/latest/USD"
RATE_DEFINITION = "USD per 1 unit of source currency"
API_CURRENCY_ALIASES = {
    "UF": "CLF",
}


_NULL_CURRENCY_STRINGS = frozenset({"NAN", "NONE", "N/A", "NA", "NAT", "NULL", ""})


def _normalize_currency(value: object) -> str | None:
    if value is None:
        return None
    # pandas NA, float NaN
    if isinstance(value, float) and value != value:
        return None
    text = str(value).strip().upper()
    if text in _NULL_CURRENCY_STRINGS:
        return None
    return text


def _load_cached_rates(cache_path: Path, as_of_date: str) -> dict[str, float]:
    if not cache_path.exists():
        return {}

    with cache_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    if payload.get("as_of_date") != as_of_date:
        return {}

    rates = payload.get("rates", {})
    if not isinstance(rates, dict):
        return {}

    cleaned: dict[str, float] = {}
    for currency, rate in rates.items():
        normalized = _normalize_currency(currency)
        if normalized is None:
            continue
        cleaned[normalized] = float(rate)
    return cleaned


def _write_cached_rates(cache_path: Path, as_of_date: str, rates: dict[str, float]) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "as_of_date": as_of_date,
        "base": "USD",
        "source": SOURCE_URL,
        "rate_definition": RATE_DEFINITION,
        "fetched_at_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "rates": {currency: rates[currency] for currency in sorted(rates)},
    }
    with cache_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, sort_keys=True)
        handle.write("\n")


def _fetch_usd_base_rates() -> dict[str, float]:
    request = urllib.request.Request(SOURCE_URL, headers={"User-Agent": "lundin-cleansing-servicios/1.0"})
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except Exception as exc:  # noqa: BLE001 - keep the operational error explicit for CLI users.
        raise RuntimeError(f"No se pudo consultar la tasa USD desde {SOURCE_URL}: {exc}") from exc

    rates = payload.get("rates")
    if not isinstance(rates, dict):
        raise RuntimeError(f"La respuesta de tasas USD no contiene 'rates': {SOURCE_URL}")

    return {str(currency).upper(): float(rate) for currency, rate in rates.items()}


def _convert_usd_base_to_usd_per_currency(currencies: Iterable[str]) -> dict[str, float]:
    usd_base_rates = _fetch_usd_base_rates()
    converted: dict[str, float] = {}
    missing: list[str] = []

    for currency in currencies:
        if currency == "USD":
            converted[currency] = 1.0
            continue

        lookup_currency = API_CURRENCY_ALIASES.get(currency, currency)
        currency_per_usd = usd_base_rates.get(lookup_currency)
        if currency_per_usd is None or currency_per_usd == 0:
            missing.append(currency)
            continue
        converted[currency] = 1.0 / float(currency_per_usd)

    if missing:
        raise RuntimeError(
            "No se encontraron tasas USD para las monedas: "
            + ", ".join(sorted(missing))
            + f". Agrega estas tasas en {SOURCE_URL} o en el cache local."
        )

    return converted


def load_usd_rates(currencies: Iterable[object], cache_path: Path, as_of_date: str) -> dict[str, float]:
    required = sorted({currency for currency in (_normalize_currency(value) for value in currencies) if currency})
    if not required:
        return {}

    cached_rates = _load_cached_rates(cache_path, as_of_date)
    cached_rates["USD"] = 1.0

    missing = [currency for currency in required if currency not in cached_rates]
    if missing:
        fetched_rates = _convert_usd_base_to_usd_per_currency(missing)
        cached_rates.update(fetched_rates)
        _write_cached_rates(cache_path, as_of_date, cached_rates)
    elif not cache_path.exists():
        _write_cached_rates(cache_path, as_of_date, cached_rates)

    return {currency: cached_rates[currency] for currency in required}