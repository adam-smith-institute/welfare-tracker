"""Build a single self-contained HTML file for sharing by email, Teams or USB.

    python3 build_standalone.py

index.html is the published artifact's source: it loads data.js, geo.js and the logo
from sibling files and d3 and Josefin Sans from CDNs. Opened on its own (an email
attachment, or a file double-clicked inside a zip) those are missing and the page is blank.

This script inlines everything - data, boundaries, d3, logo and fonts - and adds the
document shell (doctype, UTF-8 charset, viewport) the artifact host normally supplies.
Output: ../England-Welfare-Tracker-2026.html. Works offline in any modern browser.
"""
import base64, os

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "England-Welfare-Tracker-2026.html")


def read(name, mode="r"):
    with open(os.path.join(HERE, name), mode, **({} if "b" in mode else {"encoding": "utf-8"})) as f:
        return f.read()


def inline_script(js):
    return "<script>\n" + js.replace("</script", "<\\/script") + "\n</script>"


def swap(html, old, new):
    assert old in html, "not found: " + old[:70]
    return html.replace(old, new)


page = read("index.html")

# Fonts: embed the design system's Josefin Sans files instead of calling Google Fonts.
fonts = "".join(
    "@font-face{font-family:\"Josefin Sans\";font-style:%s;font-weight:100 700;font-display:swap;"
    "src:url(data:font/ttf;base64,%s) format(\"truetype\");}\n"
    % (style, base64.b64encode(read("vendor/" + fn, "rb")).decode())
    for fn, style in (("JosefinSans.ttf", "normal"), ("JosefinSans-Italic.ttf", "italic")))
for line in (
    '<link rel="preconnect" href="https://fonts.googleapis.com">\n',
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n',
    '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Josefin+Sans:ital,wght@0,300;0,400;0,600;1,300;1,400&display=swap">\n',
):
    page = swap(page, line, "")
page = swap(page, "<style>\n", "<style>\n" + fonts)

# Library, data and boundaries.
page = swap(page, '<script src="https://cdnjs.cloudflare.com/ajax/libs/d3/7.9.0/d3.min.js"></script>', inline_script(read("vendor/d3.min.js")))
page = swap(page, '<script src="data.js"></script>', inline_script(read("data.js")))
page = swap(page, '<script src="geo.js"></script>', inline_script(read("geo.js")))

# Logo as a data URI.
logo = "data:image/svg+xml;base64," + base64.b64encode(read("asi_logo_white.svg", "rb")).decode()
page = swap(page, 'src="asi_logo_white.svg"', 'src="' + logo + '"')

# Document shell: the artifact host adds this when publishing; a local file needs its own.
head_end = page.index('<svg width="0" height="0"')
doc = (
    "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n"
    "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n"
    + page[:head_end] + "</head>\n<body>\n" + page[head_end:] + "\n</body>\n</html>\n"
)

with open(OUT, "w", encoding="utf-8") as f:
    f.write(doc)
print("wrote", os.path.normpath(OUT), f"({len(doc.encode('utf-8')) / 1e6:.2f} MB)")
