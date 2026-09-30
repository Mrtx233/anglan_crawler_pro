"""name 字段：来源品牌替换、符号清理和英文单词大小写规范。"""
import re
import unicodedata
from dataclasses import dataclass


class NameValidationError(ValueError):
    pass


class UnsupportedNameCharacters(NameValidationError):
    """商品名称含不支持的文字，合并时跳过该商品。"""


# 只处理能够明确识别的域名结构；未知多段后缀交给人工确认。
COMPOUND_SUFFIXES = frozenset({
    'co.uk', 'org.uk', 'me.uk', 'com.au', 'net.au', 'org.au',
    'co.nz', 'co.jp', 'co.kr', 'co.za', 'com.cn', 'net.cn', 'org.cn',
    'com.hk', 'com.tw', 'com.sg', 'com.br', 'com.mx', 'co.in',
})
_LABEL = r'[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?'
_DOMAIN = re.compile(rf'(?<![A-Za-z0-9.-])(?:{_LABEL}\.)+[A-Za-z]{{2,63}}(?![A-Za-z0-9.-])')


@dataclass(frozen=True)
class Brand:
    domain: str
    word: str


def extract_brand(label):
    """从文件夹末级名称或去除 .xlsx 的文件名中提取唯一域名。"""
    domains = {m.group().lower().removeprefix('www.') for m in _DOMAIN.finditer(label)}
    if len(domains) != 1:
        raise NameValidationError(f'无法识别唯一域名: {label!r}，请在名称中保留一个明确域名')
    domain = domains.pop()
    parts = domain.split('.')
    suffix_parts = 2 if '.'.join(parts[-2:]) in COMPOUND_SUFFIXES else 1
    if len(parts) != suffix_parts + 1:
        raise NameValidationError(f'域名包含子域名或未配置的多段后缀，需确认品牌: {domain}')
    return Brand(domain, parts[0])


def normalize_name(value, old_brand, new_brand, *, whitelist=('USB',)):
    """先替换完整域名/品牌词，再清理符号；不匹配其他单词内部。"""
    if not isinstance(value, str):
        raise NameValidationError('name 必须是字符串')
    # 完整域名优先，避免留下 co uk 等后缀；前后不允许字母数字或下划线。
    domain_pattern = (r'(?<![A-Za-z0-9_])(?:https?://)?(?:www\.)?'
                      + re.escape(old_brand.domain) + r'(?![A-Za-z0-9_]|\.[A-Za-z0-9])')
    text = re.sub(domain_pattern, lambda m: new_brand.word, value, flags=re.IGNORECASE)
    text = re.sub(r'(?<![A-Za-z0-9_])' + re.escape(old_brand.word)
                  + r'(?![A-Za-z0-9_])', lambda m: new_brand.word, text, flags=re.IGNORECASE)
    # 统一独立数字分数；不对尚未约定的 Unicode 分数字符擅自改写。
    text = re.sub(r'(?<![\w/])(\d+) */ *(\d+)(?![\w/])', r'\1/\2', text)
    fractions = {m.start() + m.group().index('/')
                 for m in re.finditer(r'(?<![\w/])[0-9]+/[0-9]+(?![\w/])', text)}
    cleaned = []
    for index, char in enumerate(text):
        category = unicodedata.category(char)
        if char == '/' and index in fractions:
            cleaned.append(char)
        elif char.isspace():
            cleaned.append(' ')
        elif category.startswith('C') or char in ('\ufe0e', '\ufe0f', '\u20e3'):
            continue
        elif category.startswith(('P', 'S')):
            cleaned.append(' ')
        else:
            cleaned.append(char)
    text = re.sub(' +', ' ', ''.join(cleaned)).strip()
    if not text:
        raise NameValidationError('name 清理后不能为空')
    if not re.fullmatch(r'[A-Za-z0-9 /]+', text):
        raise UnsupportedNameCharacters('name 含非英文文字或未支持字符')
    canonical = {word.lower(): word for word in whitelist}
    words = []
    for word in text.split(' '):
        if '/' in word:
            if not re.fullmatch(r'[0-9]+/[0-9]+', word):
                raise NameValidationError('斜杠仅允许用于独立数字分数，如 3/4')
            if not word.split('/')[1].strip('0'):
                raise NameValidationError('分数分母不能为零')
            words.append(word)
        else:
            words.append(canonical.get(word.lower(), re.sub(
                '[a-z]', lambda m: m.group().upper(), word.lower(), count=1)))
    return ' '.join(words)
