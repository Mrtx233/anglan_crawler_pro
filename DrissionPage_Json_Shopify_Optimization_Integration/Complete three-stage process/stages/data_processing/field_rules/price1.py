"""price1：有限、非负数字，使用十进制四舍五入保留两位小数。"""
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


class PriceValidationError(ValueError):
    pass


def normalize_price1(value):
    """校验并规范 price1：有限非负数字，十进制四舍五入到两位小数。

    返回 Decimal（避免二进制浮点误差），因此能与 price2 做精确比较。
    布尔值被显式拒绝——True/False 在 Python 里是 int 的子类，
    放过会得到 1.00/0.00 这种静默错误。
    """
    if isinstance(value, bool) or not isinstance(value, (str, int, float, Decimal)):
        raise PriceValidationError('价格必须是数字或数字字符串')
    try:
        amount = Decimal(str(value).strip())
        if not amount.is_finite() or amount < 0:
            raise PriceValidationError('价格必须是有限的非负数字')
        return amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    except InvalidOperation as exc:
        raise PriceValidationError('价格为空、不是有效数字或超出可处理范围') from exc
