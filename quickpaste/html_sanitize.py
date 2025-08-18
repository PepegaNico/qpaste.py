import re


ALLOWED_TAGS = {
    'b', 'i', 'u', 'em', 'strong', 'a', 'br', 'p', 'span'
}
ALLOWED_ATTRS = {
    'a': {'href'},
    'span': {'style'},
}


def sanitize_html(html: str) -> str:
    # very light-weight sanitizer for the current scope
    # remove script/style blocks completely
    html = re.sub(r'<\s*(script|style)[^>]*>[\s\S]*?<\s*/\s*\1\s*>', '', html, flags=re.IGNORECASE)
    # remove on* event attributes
    html = re.sub(r'\s+on\w+\s*=\s*(["\']).*?\1', '', html, flags=re.IGNORECASE)
    # drop tags not in whitelist
    def repl_tag(match):
        tag = match.group(1)
        name = match.group(2).lower()
        rest = match.group(3) or ''
        if name not in ALLOWED_TAGS:
            return ''
        # filter attributes
        if name in ALLOWED_ATTRS:
            allowed = ALLOWED_ATTRS[name]
            attrs = re.findall(r'(\w+)\s*=\s*(["\']).*?\2', rest)
            keep = [f'{k}={v}' for k, v in attrs if k in allowed]
            rest = (' ' + ' '.join(keep)) if keep else ''
        else:
            rest = ''
        return f'<{tag}{name}{rest}>'
    html = re.sub(r'<\s*(/?)\s*([a-zA-Z0-9]+)([^>]*)>', repl_tag, html)
    return html

