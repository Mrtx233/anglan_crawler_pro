from __future__ import annotations

import json
import urllib.request

_exchange_rates_cache = {}


def fetch_exchange_rates(log=print):
    global _exchange_rates_cache
    if _exchange_rates_cache:
        return _exchange_rates_cache
    try:
        url = "https://open.er-api.com/v6/latest/USD"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read().decode())
        _exchange_rates_cache = data.get("rates", {})
        log(f"[汇率] 成功获取 {_exchange_rates_cache.get('USD', 1)} 相关汇率")
    except Exception as e:
        log(f"[汇率] 获取失败: {e}")
    return _exchange_rates_cache


def convert_price_to_usd(price_str, currency, rates=None):
    if not price_str or not currency:
        return price_str
    if rates is None:
        rates = _exchange_rates_cache
    if not rates:
        return price_str
    currency = str(currency).upper().strip()
    if currency == "USD":
        return price_str
    rate = rates.get(currency)
    if not rate or rate == 0:
        return price_str
    try:
        amount = float(str(price_str).strip())
    except (ValueError, TypeError):
        return price_str
    usd_amount = amount / rate
    return f"{usd_amount:.2f}"


def format_price(value, rates=None, currency=None):
    if value in (None, ""):
        return ""
    raw = str(value).strip()
    if not raw:
        return ""
    if rates and currency:
        raw = convert_price_to_usd(raw, currency, rates)
    if "." in raw:
        try:
            amount = float(raw)
        except ValueError:
            return raw
    else:
        try:
            amount = int(raw) / 100
        except ValueError:
            return raw
    return str(int(amount)) if amount.is_integer() else f"{amount:.2f}".rstrip("0").rstrip(".")


def _is_zero_price(value):
    try:
        return float(str(value).strip()) == 0
    except (ValueError, TypeError):
        return False


def format_price2(variant, rates=None):
    currency = variant.get("price_currency", "")
    compare_at_price = variant.get("compare_at_price")
    if compare_at_price and str(compare_at_price).strip() and not _is_zero_price(compare_at_price):
        return format_price(compare_at_price, rates=rates, currency=currency)
    return format_price(variant.get("price", ""), rates=rates, currency=currency)
