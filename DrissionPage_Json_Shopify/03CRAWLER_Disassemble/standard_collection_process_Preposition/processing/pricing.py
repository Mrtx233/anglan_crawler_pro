from __future__ import annotations

from bisect import bisect_left, bisect_right

import pandas as pd

from ..resources.price_library import PRICE_LIBRARY
from .cells import _clean_cell
from .style_codec import _style_split, _style_split_first


# ============================================================
# 价格匹配配置
# ============================================================

PRICE_LIBRARY_SORTED = sorted(PRICE_LIBRARY)

# 根据商品原价在整个文件中的价格百分位，决定允许匹配的倍率范围。
#
# 格式：
# (
#     百分位上限,
#     最低倍率,
#     最高倍率,
# )
#
# 例如：
# 原价 39，并且属于整个文件最低价格：
# 39 × 0.20 = 7.8
# 39 × 0.80 = 31.2
# 因此 8.99、9.xx、10.xx 等低价可以进入候选池。
PRICE_RATIO_RULES = (
    (0.10, 0.20, 0.80),
    (0.30, 0.40, 0.90),
    (0.70, 0.60, 1.10),
    (0.90, 0.70, 1.30),
    (1.00, 0.80, 1.50),
)

# 正常倍率范围完全找不到价格时，逐步扩大候选区间。
FALLBACK_EXPANSIONS = (0.10, 0.20)

# 最终兜底时，从目标价附近选多少价格组成候选池。
# 避免高价商品因为“优先使用未使用价格”而突然匹配到 8.99。
FALLBACK_POOL_MIN_SIZE = 12
FALLBACK_POOL_MULTIPLIER = 4


# ============================================================
# 基础工具
# ============================================================

def _to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _split_variant(segment):
    """拆出一个 styles1 变体段。

    opts 保留转义序列原样，重组时直接拼回，避免二次转义。
    """
    text, image, has_image = _style_split_first(_clean_cell(segment), "@")
    parts = text.rsplit("$", 2)
    opts, price1, price2 = parts if len(parts) == 3 else (text, "", "")

    return {
        "raw": segment,
        "opts": opts,
        "price1": price1,
        "price2": price2,
        "suffix": f"@{image}" if has_image else "",
    }


def _parse_styles(raw):
    """解析 styles1，返回 (Option1 名称, [变体 dict, ...])。"""
    text = _clean_cell(raw)

    if not text:
        return "", []

    segments = _style_split(text, "#")

    return (
        segments[0],
        [_split_variant(segment) for segment in segments[1:]],
    )


def _get_original_price(row):
    """获取一行商品用于匹配的原始基准价格。

    优先 price2；
    price2 无效时使用 price1。
    """
    price2 = _to_float(row.get("price2"))

    if price2 is not None and price2 > 0:
        return price2

    price1 = _to_float(row.get("price1"))

    if price1 is not None and price1 > 0:
        return price1

    return None


# ============================================================
# 原价修正
# ============================================================

def fix_original_prices(source_df):
    """原价修正：

    price2 < price1 → price2 = price1

    styles1 内嵌价：
    p2 < p1 → p2 = p1

    返回：
    (
        修正后的表,
        price1/price2 修正行数,
        styles1 内嵌修正变体数
    )
    """
    df = source_df.copy(deep=True).reset_index(drop=True)

    for col in ("price1", "price2"):
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col],
                errors="coerce",
            ).astype("float64")

    row_fixed = 0
    segment_fixed = 0

    for index, row in df.iterrows():
        price1 = _to_float(row.get("price1"))
        price2 = _to_float(row.get("price2"))

        if (
            price1 is not None
            and price2 is not None
            and price2 < price1
        ):
            df.at[index, "price2"] = price1
            row_fixed += 1

        option1_name, variants = _parse_styles(
            row.get("styles1")
        )

        if not variants:
            continue

        rebuilt = [option1_name]
        changed = 0

        for variant in variants:
            inner1 = _to_float(variant["price1"])
            inner2 = _to_float(variant["price2"])

            if (
                inner1 is not None
                and inner2 is not None
                and inner2 < inner1
            ):
                # 沿用原始文本而不是 float，
                # 避免 20.00 被改成 20.0。
                price_text = variant["price1"]

                rebuilt.append(
                    f"{variant['opts']}"
                    f"${price_text}"
                    f"${price_text}"
                    f"{variant['suffix']}"
                )

                changed += 1
            else:
                rebuilt.append(variant["raw"])

        if changed:
            df.at[index, "styles1"] = "#".join(rebuilt)
            segment_fixed += changed

    return df, row_fixed, segment_fixed


# ============================================================
# 百分位计算
# ============================================================

def _build_match_rows(df):
    """预解析所有真正需要匹配价格的行。

    避免 match_prices 中重复解析 styles1。
    """
    rows = {}

    for index, row in df.iterrows():
        option1_name, variants = _parse_styles(
            row.get("styles1")
        )

        if not variants:
            continue

        rows[index] = {
            "option1_name": option1_name,
            "variants": variants,
            "original": _get_original_price(row),
        }

    return rows


def _get_price_percentiles(match_rows):
    """计算商品原价的全局百分位。

    使用“唯一价格等级”计算，而不是简单按照行数 rank。

    好处：
    如果很多商品价格都是 39，
    不会因为 39 出现次数太多，
    导致最低价格的 percentile 被抬到 0.2、0.3。

    最低唯一价格始终接近 0；
    最高唯一价格始终接近 1；
    相同原价拥有完全相同 percentile。
    """
    valid_prices = {
        index: info["original"]
        for index, info in match_rows.items()
        if info["original"] is not None
        and info["original"] > 0
    }

    if not valid_prices:
        return {}

    unique_prices = sorted(set(valid_prices.values()))

    # 整个文件只有一种原价时，没有真正意义上的高低排名。
    # 放在中间位置更安全，避免全部商品进入最低档或最高档。
    if len(unique_prices) == 1:
        return {
            index: 0.5
            for index in valid_prices
        }

    denominator = len(unique_prices) - 1

    percentile_by_price = {
        price: position / denominator
        for position, price in enumerate(unique_prices)
    }

    return {
        index: percentile_by_price[price]
        for index, price in valid_prices.items()
    }


# ============================================================
# 价格区间与目标价格
# ============================================================

def _get_ratio_range(percentile):
    """根据全局价格百分位获取合理倍率范围。"""
    percentile = max(0.0, min(1.0, percentile))

    for upper_percentile, lower_ratio, upper_ratio in PRICE_RATIO_RULES:
        if percentile <= upper_percentile:
            return lower_ratio, upper_ratio

    # 理论上不会到这里。
    _, lower_ratio, upper_ratio = PRICE_RATIO_RULES[-1]

    return lower_ratio, upper_ratio


def _get_target_price(
    original,
    percentile,
    lower_bound,
    upper_bound,
):
    """计算理想目标价格。

    percentile 越低：
    目标越靠近候选区间低端。

    percentile 越高：
    目标越靠近候选区间高端。

    例如：

    percentile = 0
    → target 接近 lower_bound

    percentile = 0.5
    → target 接近区间中间

    percentile = 1
    → target 接近 upper_bound
    """
    del original  # 保留参数语义，当前公式无需直接使用。

    percentile = max(0.0, min(1.0, percentile))

    return (
        lower_bound
        + (upper_bound - lower_bound) * percentile
    )


def _prices_in_range(lower_bound, upper_bound):
    """利用已排序价格库快速获取指定区间价格。"""
    if not PRICE_LIBRARY_SORTED:
        return []

    start = bisect_left(
        PRICE_LIBRARY_SORTED,
        lower_bound,
    )

    end = bisect_right(
        PRICE_LIBRARY_SORTED,
        upper_bound,
    )

    return PRICE_LIBRARY_SORTED[start:end]


def _normalized_distance(price, target):
    """候选价与目标价的归一化距离。

    使用比例距离而不是绝对金额距离，
    避免高价格天然产生更大的 distance。
    """
    denominator = max(abs(target), 1.0)

    return abs(price - target) / denominator


# ============================================================
# 候选池
# ============================================================

def _get_candidate_pool(
    lower_bound,
    upper_bound,
    target_price,
    variant_count,
):
    """获取当前商品的合理候选价格。

    第一层：
    正常倍率区间。

    第二层：
    候选完全为空时逐步扩大范围。

    第三层：
    仍然为空时，从目标价格附近构造小型候选池。

    第三层尤其用于：
    original > 价格库最大值
    等特殊情况。
    """
    candidates = _prices_in_range(
        lower_bound,
        upper_bound,
    )

    if candidates:
        return candidates

    # 第一层 fallback：
    # 逐步扩大合理价格区间。
    for expansion in FALLBACK_EXPANSIONS:
        expanded_lower = max(
            0.0,
            lower_bound * (1.0 - expansion),
        )

        expanded_upper = (
            upper_bound * (1.0 + expansion)
        )

        candidates = _prices_in_range(
            expanded_lower,
            expanded_upper,
        )

        if candidates:
            return candidates

    # 第二层 fallback：
    # 不直接从全部价格中按照 usage 选择，
    # 否则一个 $800 商品可能因为 8.99 从未使用，
    # 错误地匹配成 8.99。
    #
    # 先按目标价距离取附近的一小块，
    # 再在这块里面做 usage 均衡。
    if not PRICE_LIBRARY_SORTED:
        return []

    pool_size = max(
        FALLBACK_POOL_MIN_SIZE,
        variant_count * FALLBACK_POOL_MULTIPLIER,
    )

    pool_size = min(
        pool_size,
        len(PRICE_LIBRARY_SORTED),
    )

    nearest = sorted(
        PRICE_LIBRARY_SORTED,
        key=lambda price: (
            _normalized_distance(
                price,
                target_price,
            ),
            price,
        ),
    )

    return nearest[:pool_size]


# ============================================================
# Variant 分配
# ============================================================

def _candidate_score(
    price,
    target_price,
    price_usage,
):
    """候选价评分。

    第一优先级：
    使用次数少。

    第二优先级：
    离目标价近。

    第三优先级：
    价格较小。

    因为候选价格已经经过“合理区间”过滤，
    所以这里可以让覆盖率优先于距离。
    """
    return (
        price_usage.get(price, 0),
        _normalized_distance(
            price,
            target_price,
        ),
        price,
    )


def _assign_variant_prices(
    candidates,
    variant_count,
    target_price,
    price_usage,
):
    """给一个商品的多个 Variant 分配价格。

    规则：

    1. 尽量不同
    2. 优先使用次数少的价格
    3. 相同 usage 时优先接近 target
    4. 候选不足时允许重复
    5. 重复时尽量均衡
    6. 最终从低到高排序
    """
    if variant_count <= 0:
        return []

    if not candidates:
        return [None] * variant_count

    # 去掉候选池内可能出现的重复值。
    unique_candidates = list(
        dict.fromkeys(candidates)
    )

    if not unique_candidates:
        return [None] * variant_count

    # --------------------------------------------------------
    # 候选数量足够：
    # 尽量每个 Variant 使用不同价格。
    # --------------------------------------------------------
    if len(unique_candidates) >= variant_count:
        ranked = sorted(
            unique_candidates,
            key=lambda price: _candidate_score(
                price,
                target_price,
                price_usage,
            ),
        )

        selected = ranked[:variant_count]

        # Variant 最终价格必须递增或不下降。
        return sorted(selected)

    # --------------------------------------------------------
    # 候选数量不足：
    # 所有不同价格至少使用一次，
    # 剩余 Variant 再做均衡重复。
    # --------------------------------------------------------

    selected = list(unique_candidates)

    # 当前商品内部每个价格已经使用一次。
    local_usage = {
        price: 1
        for price in unique_candidates
    }

    while len(selected) < variant_count:

        # 优先选择：
        #
        # 全局 usage + 当前商品内部 usage
        # 最低的价格。
        #
        # 这样例如：
        #
        # 3 个候选
        # 6 个 Variant
        #
        # 更容易得到：
        #
        # 10 10
        # 20 20
        # 30 30
        #
        # 而不是：
        #
        # 10 10 10 10
        # 20
        # 30
        best = min(
            unique_candidates,
            key=lambda price: (
                price_usage.get(price, 0)
                + local_usage.get(price, 0),
                _normalized_distance(
                    price,
                    target_price,
                ),
                price,
            ),
        )

        selected.append(best)

        local_usage[best] = (
            local_usage.get(best, 0) + 1
        )

    # 重新排序后对应原 styles1 Variant 顺序。
    return sorted(selected)


# ============================================================
# 核心价格匹配
# ============================================================

def match_prices(source_df):
    """执行全局价格匹配。

    新算法：

    1. 获取所有待匹配商品原价
    2. 计算全局价格 percentile
    3. 按 percentile 设置合理倍率区间
    4. 根据 percentile 计算目标价格
    5. 构造候选价格池
    6. 优先使用 price_usage 较低的价格
    7. 相同 usage 时选择更接近目标价的价格
    8. 同商品 Variant 尽量使用不同价格
    9. 候选不足时均衡重复
    10. Variant 最终按照价格递增分配

    返回：

    (
        匹配后的表,
        未使用价格列表,
        {
            行号: [
                每个 Variant 的匹配价格
            ]
        }
    )
    """
    df = source_df.copy(
        deep=True
    ).reset_index(drop=True)

    for col in ("price1", "price2"):
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col],
                errors="coerce",
            ).astype("float64")

    # 如果价格库为空，不进行任何匹配。
    if not PRICE_LIBRARY_SORTED:
        return df, [], {}

    # --------------------------------------------------------
    # 1. 预解析所有待匹配商品
    # --------------------------------------------------------

    match_rows = _build_match_rows(df)

    # --------------------------------------------------------
    # 2. 计算全局价格百分位
    # --------------------------------------------------------

    price_percentiles = _get_price_percentiles(
        match_rows
    )

    # --------------------------------------------------------
    # 3. 记录整个文件每个价格的使用次数
    # --------------------------------------------------------

    price_usage = {
        price: 0
        for price in PRICE_LIBRARY_SORTED
    }

    matched_prices = {}

    # --------------------------------------------------------
    # 4. 按原始 DataFrame 顺序处理
    # --------------------------------------------------------

    for index in range(len(df)):

        info = match_rows.get(index)

        if info is None:
            continue

        option1_name = info["option1_name"]
        variants = info["variants"]
        original = info["original"]

        variant_count = len(variants)

        if (
            original is None
            or original <= 0
            or variant_count == 0
        ):
            continue

        percentile = price_percentiles.get(
            index,
            0.5,
        )

        # ----------------------------------------------------
        # 当前价格层级允许的倍率
        # ----------------------------------------------------

        lower_ratio, upper_ratio = (
            _get_ratio_range(percentile)
        )

        lower_bound = (
            original * lower_ratio
        )

        upper_bound = (
            original * upper_ratio
        )

        # ----------------------------------------------------
        # 理想目标价格
        # ----------------------------------------------------

        target_price = _get_target_price(
            original=original,
            percentile=percentile,
            lower_bound=lower_bound,
            upper_bound=upper_bound,
        )

        # ----------------------------------------------------
        # 构造候选价格池
        # ----------------------------------------------------

        candidates = _get_candidate_pool(
            lower_bound=lower_bound,
            upper_bound=upper_bound,
            target_price=target_price,
            variant_count=variant_count,
        )

        # ----------------------------------------------------
        # 给多个 Variant 分配价格
        # ----------------------------------------------------

        assigned = _assign_variant_prices(
            candidates=candidates,
            variant_count=variant_count,
            target_price=target_price,
            price_usage=price_usage,
        )

        used = [
            price
            for price in assigned
            if price is not None
        ]

        if not used:
            continue

        # ----------------------------------------------------
        # 更新全局使用次数
        # ----------------------------------------------------

        for price in used:
            price_usage[price] = (
                price_usage.get(price, 0) + 1
            )

        # ----------------------------------------------------
        # 行级 price1 / price2
        # ----------------------------------------------------

        matched_low = min(used)
        matched_original = max(used)

        df.at[index, "price1"] = matched_low
        df.at[index, "price2"] = matched_original

        # ----------------------------------------------------
        # 重建 styles1
        #
        # 每个 Variant：
        #
        # price1 = 自己匹配的价格
        # price2 = 当前商品最高匹配价
        # ----------------------------------------------------

        rebuilt = [option1_name]

        for variant, matched in zip(
            variants,
            assigned,
        ):
            if matched is None:
                rebuilt.append(
                    variant["raw"]
                )
                continue

            rebuilt.append(
                f"{variant['opts']}"
                f"${matched}"
                f"${matched_original}"
                f"{variant['suffix']}"
            )

        df.at[index, "styles1"] = "#".join(
            rebuilt
        )

        matched_prices[index] = assigned

    # --------------------------------------------------------
    # 5. 返回整个任务结束后一次都没用过的价格
    # --------------------------------------------------------

    unused_prices = [
        price
        for price in PRICE_LIBRARY_SORTED
        if price_usage.get(price, 0) == 0
    ]

    return (
        df,
        unused_prices,
        matched_prices,
    )