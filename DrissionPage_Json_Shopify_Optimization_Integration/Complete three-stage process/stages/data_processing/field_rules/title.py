"""title 规范：英文 Title Case、严格字符检查与类目路径格式。"""
import re
import unicodedata
from dataclasses import dataclass


class TitleValidationError(ValueError):
    pass


@dataclass(frozen=True)
class TitleRules:
    max_length: int = 100
    max_levels: int = 2
    allow_digits: bool = True
    # 大小写不敏感匹配单词，值为输出的标准拼写。
    whitelist: tuple[str, ...] = ('USB',)


DEFAULT_RULES = TitleRules()


def normalize_title(value, rules=DEFAULT_RULES):
    """清理符号并规范空格/大小写；空层级和超限内容拒绝写入。"""
    if not isinstance(value, str):
        raise TitleValidationError('title 必须是字符串')
    # 分数先规范并保护斜杠；其他标点用空格隔开，避免两个词粘连。
    text = re.sub(r'(?<![\w/])(\d+) */ *(\d+)(?![\w/])', r'\1/\2', value)
    fraction_slashes = {
        match.start() + match.group().index('/')
        for match in re.finditer(r'(?<![\w/])[0-9]+/[0-9]+(?![\w/])', text)
    }
    cleaned = []
    for index, char in enumerate(text):
        category = unicodedata.category(char)
        if char == '>' or (char == '/' and index in fraction_slashes):
            cleaned.append(char)
        elif char.isspace():
            cleaned.append(' ')
        elif category.startswith('C') or char in ('\ufe0e', '\ufe0f', '\u20e3'):
            # 隐藏字符及 emoji 的展示控制字符直接删除。
            continue
        elif category.startswith(('P', 'S')):
            cleaned.append(' ')
        else:
            cleaned.append(char)
    text = ''.join(cleaned).strip(' ')
    if not text:
        raise TitleValidationError('title 清理后不能为空')
    if any(ord(char) > 126 or ord(char) < 32 for char in text):
        raise TitleValidationError('非半角或非英文文字需人工确认')
    if not re.fullmatch(r'[A-Za-z0-9 >/]+', text):
        raise TitleValidationError('包含禁止符号')
    if not rules.allow_digits and re.search(r'\d', text):
        raise TitleValidationError('当前规则不允许数字')
    levels = text.split('>')
    if len(levels) > rules.max_levels:
        raise TitleValidationError(f'分类最多允许 {rules.max_levels} 级')
    canonical = {word.lower(): word for word in rules.whitelist}
    normalized = []
    for level in levels:
        level = re.sub(r' +', ' ', level.strip(' '))
        if not level:
            raise TitleValidationError('分类不能包含空层级')
        level = re.sub(r'(?<![A-Za-z0-9/])(\d+) */ *(\d+)(?![A-Za-z0-9/])', r'\1/\2', level)
        words = []
        for word in level.split(' '):
            if '/' in word:
                if not re.fullmatch(r'[0-9]+/[0-9]+', word):
                    raise TitleValidationError('斜杠仅允许用于独立数字分数，如 3/4')
                if not word.split('/')[1].strip('0'):
                    raise TitleValidationError('分数分母不能为零')
                words.append(word)
            elif word.lower() in canonical:
                words.append(canonical[word.lower()])
            else:
                lowered = word.lower()
                # 数字允许出现在词首，如 3d -> 3D。
                words.append(re.sub(r'[a-z]', lambda m: m.group().upper(), lowered, count=1))
        normalized.append(' '.join(words))
    result = ' > '.join(normalized)
    if len(result) > rules.max_length:
        raise TitleValidationError(f'title 长度 {len(result)} 超过 {rules.max_length}，需业务确认')
    return result
