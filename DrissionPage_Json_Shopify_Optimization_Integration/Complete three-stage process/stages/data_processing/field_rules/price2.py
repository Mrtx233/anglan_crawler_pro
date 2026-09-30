"""price2：两位小数；低于 price1 时改为 price1 的 1.1 倍。"""
from decimal import Decimal, ROUND_HALF_UP

from stages.data_processing.field_rules.price1 import normalize_price1


def normalize_price2(value, price1):
    """规范 price2：两位小数；低于 price1 时替换为 price1 的 1.1 倍。

    两个入参都先过 normalize_price1 校验，非法值会抛 PriceValidationError。
    """
    minimum = normalize_price1(price1)
    amount = normalize_price1(value)
    if amount < minimum:
        amount = (minimum * Decimal('1.1')).quantize(
            Decimal('0.01'), rounding=ROUND_HALF_UP)
    return amount
