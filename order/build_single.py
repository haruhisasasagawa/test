"""
order/ の HTML・CSS・JS を1つのHTMLファイルにまとめる。

    python3 order/build_single.py [出力パス]

配布は1ファイルで済ませたいので（メールで送る、共有フォルダに置く）、
開発は分割したまま、配布物だけ結合する。
"""

import re
import sys
from pathlib import Path

BASE = Path(__file__).parent


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / 'dist' / '売店発注ツール.html'
    html = (BASE / 'index.html').read_text(encoding='utf-8')
    css = (BASE / 'style.css').read_text(encoding='utf-8')
    js = '\n'.join((BASE / name).read_text(encoding='utf-8') for name in ('seed.js', 'app.js'))

    html = html.replace('<link rel="stylesheet" href="style.css">',
                        '<style>\n' + css + '\n</style>')
    # 置換文字列としてではなく関数で返す。JS中の \n や \1 が
    # エスケープとして解釈されてしまうのを防ぐ。
    script = '\n<script>\n' + js + '\n</script>'
    html = re.sub(r'\s*<script src="seed\.js"></script>\s*<script src="app\.js"></script>',
                  lambda _: script, html)

    if 'href="style.css"' in html or 'src="app.js"' in html:
        raise SystemExit('結合に失敗しました。index.html の読み込みタグを確認してください。')

    # --artifact: 外側のひな形（doctype/html/head/body）を持たない断片を出力する。
    # 公開先が自前のひな形で包むため、title・style・本文・scriptだけを渡す。
    if '--artifact' in sys.argv:
        title = re.search(r'<title>(.*?)</title>', html, re.S).group(1)
        style = re.search(r'<style>.*?</style>', html, re.S).group(0)
        body = re.search(r'<body>(.*)</body>', html, re.S).group(1)
        html = f'<title>{title}</title>\n{style}\n{body}'

    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding='utf-8')
    print(f'saved: {out}  ({len(html.encode("utf-8")) / 1024:.0f} KB)')


if __name__ == '__main__':
    main()
