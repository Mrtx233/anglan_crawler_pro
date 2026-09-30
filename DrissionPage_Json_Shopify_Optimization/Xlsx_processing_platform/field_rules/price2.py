"""price2：两位小数；低于 price1 时改为 price1 的 1.1 倍。"""
from decimal import Decimal, ROUND_HALF_UP

from field_rules.price1 import normalize_price1


def normalize_price2(value, price1):
    minimum = normalize_price1(price1)
    amount = normalize_price1(value)
    if amount < minimum:
        amount = (minimum * Decimal('1.1')).quantize(
            Decimal('0.01'), rounding=ROUND_HALF_UP)
    return amount
