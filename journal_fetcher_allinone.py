from __future__ import annotations

import argparse
import difflib
import html
import json
import re
import shutil
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import zipfile
from dataclasses import dataclass
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape

CROSSREF_API_URL = "https://api.crossref.org/works"
OPENALEX_SOURCES_URL = "https://api.openalex.org/sources"
OPENALEX_WORKS_URL = "https://api.openalex.org/works"
DEFAULT_PORT = 8765
README_TEXT = """使用说明
1. 运行 `python journal_fetcher_allinone.py` 使用终端模式。
2. 运行 `python journal_fetcher_allinone.py --web` 启动网页模式。
3. 主题留空时，会抓取该时间段内该期刊的所有文章。
4. 结果默认导出到脚本同目录，格式为 xlsx。
5. 只需要 Python 3，无需额外安装第三方依赖。"""


@dataclass
class ArticleRow:
    title: str
    journal: str
    date: str
    topic: str
    abstract: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按期刊名、主题和时间范围检索文献，并导出为 Excel。")
    parser.add_argument("--journal")
    parser.add_argument("--topic", help="留空则抓取所有文章")
    parser.add_argument("--from-date")
    parser.add_argument("--until-date")
    parser.add_argument("--max-results", type=int, default=0, help="最多导出多少条；0 表示不限制")
    parser.add_argument("--output")
    parser.add_argument("--mailto")
    parser.add_argument("--web", action="store_true")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--no-browser", action="store_true")
    return parser.parse_args()


def prompt_if_missing(value: str | None, label: str, required: bool = True) -> str | None:
    if value is not None:
        return value.strip()
    while True:
        entered = input(f"{label}: ").strip()
        if entered or not required:
            return entered or None
        print(f"{label}不能为空，请重新输入。")


def validate_date(date_text: str | None, field_name: str) -> str | None:
    if not date_text:
        return None
    try:
        datetime.strptime(date_text, "%Y-%m-%d")
    except ValueError as exc:
        raise SystemExit(f"{field_name}格式错误，请使用 YYYY-MM-DD。") from exc
    return date_text


def build_params(journal: str, topic: str | None, from_date: str | None, until_date: str | None, max_results: int) -> dict[str, str]:
    filters = ["type:journal-article"]
    if from_date:
        filters.append(f"from-pub-date:{from_date}")
    if until_date:
        filters.append(f"until-pub-date:{until_date}")
    params = {
        "query.container-title": journal,
        "rows": str(max(max_results * 3, 200) if max_results else 200),
        "filter": ",".join(filters),
        "select": "title,container-title,published-print,published-online,issued,abstract,DOI",
    }
    if topic:
        params["query"] = topic
    return params


def fetch_articles(journal: str, topic: str | None, from_date: str | None, until_date: str | None, max_results: int, mailto: str | None) -> list[ArticleRow]:
    source = find_openalex_source(journal, mailto)
    if source:
        rows = fetch_articles_from_openalex(source, journal, topic, from_date, until_date, max_results, mailto)
        if rows:
            return rows

    query_string = urllib.parse.urlencode(build_params(journal, topic, from_date, until_date, max_results))
    request = urllib.request.Request(
        f"{CROSSREF_API_URL}?{query_string}",
        headers={"User-Agent": "JournalFetcher/1.0" + (f" (mailto:{mailto})" if mailto else "")},
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        payload = json.loads(response.read().decode("utf-8"))
    items = payload.get("message", {}).get("items", [])
    rows: list[ArticleRow] = []
    topic_text = (topic or "").strip() or "所有文章"
    for item in items:
        journal_name = first_text(item.get("container-title", []))
        if not journal_matches(journal, journal_name):
            continue
        title = first_text(item.get("title", [])) or "(无标题)"
        doi = (item.get("DOI") or "").strip()
        abstract = clean_abstract(item.get("abstract"))
        if not abstract:
            abstract = fetch_abstract_from_europe_pmc(title=title, journal=journal_name or journal, doi=doi, mailto=mailto)
        rows.append(
            ArticleRow(
                title=title,
                journal=journal_name or journal,
                date=extract_date(item),
                topic=topic_text,
                abstract=abstract or "(未检索到摘要)",
            )
        )
        if max_results and len(rows) >= max_results:
            break
    return rows


def request_json(url: str, params: dict[str, str], mailto: str | None) -> dict:
    headers = {}
    if mailto:
        headers["User-Agent"] = f"JournalFetcher/1.0 (mailto:{mailto})"
    request = urllib.request.Request(f"{url}?{urllib.parse.urlencode(params)}", headers=headers)
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def find_openalex_source(journal: str, mailto: str | None) -> dict | None:
    params = {
        "search": journal,
        "per-page": "10",
        "select": "id,display_name,issn,issn_l,type",
    }
    try:
        payload = request_json(OPENALEX_SOURCES_URL, params, mailto)
    except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError):
        return None

    candidates = payload.get("results", [])
    if not candidates:
        return None

    scored = sorted(
        candidates,
        key=lambda item: score_source_match(journal, item.get("display_name", "")),
        reverse=True,
    )
    best = scored[0]
    if score_source_match(journal, best.get("display_name", "")) < 0.45:
        return None
    return best


def score_source_match(user_input: str, candidate: str) -> float:
    left = normalize_text(user_input)
    right = normalize_text(candidate)
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    if left in right or right in left:
        return 0.92

    token_left = meaningful_tokens(user_input)
    token_right = meaningful_tokens(candidate)
    overlap = len(token_left & token_right)
    token_score = overlap / max(len(token_left), 1)
    similarity = difflib.SequenceMatcher(None, left, right).ratio()
    return max(similarity, token_score)


def fetch_articles_from_openalex(
    source: dict,
    journal: str,
    topic: str | None,
    from_date: str | None,
    until_date: str | None,
    max_results: int,
    mailto: str | None,
) -> list[ArticleRow]:
    filters = [f"primary_location.source.id:{source['id']}", "type:article"]
    if from_date:
        filters.append(f"from_publication_date:{from_date}")
    if until_date:
        filters.append(f"to_publication_date:{until_date}")
    if topic:
        filters.append(f"title_and_abstract.search:{topic}")

    params = {
        "filter": ",".join(filters),
        "per-page": "200",
        "page": "1",
        "sort": "publication_date:desc",
        "select": "display_name,publication_date,primary_location,abstract_inverted_index",
    }
    topic_text = (topic or "").strip() or "所有文章"
    rows: list[ArticleRow] = []

    while True:
        try:
            payload = request_json(OPENALEX_WORKS_URL, params, mailto)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError):
            break

        results = payload.get("results", [])
        if not results:
            break

        for item in results:
            primary_location = item.get("primary_location") or {}
            source_info = primary_location.get("source") or {}
            rows.append(
                ArticleRow(
                    title=(item.get("display_name") or "").strip() or "(无标题)",
                    journal=(source_info.get("display_name") or source.get("display_name") or journal).strip(),
                    date=(item.get("publication_date") or "").strip(),
                    topic=topic_text,
                    abstract=parse_openalex_abstract(item.get("abstract_inverted_index")) or "(未检索到摘要)",
                )
            )
            if max_results and len(rows) >= max_results:
                return rows

        current_page = int(params["page"])
        params["page"] = str(current_page + 1)

    return rows


def parse_openalex_abstract(abstract_index: dict | None) -> str:
    if not abstract_index:
        return ""

    positioned_words: list[tuple[int, str]] = []
    for word, positions in abstract_index.items():
        for pos in positions:
            positioned_words.append((int(pos), word))
    if not positioned_words:
        return ""

    positioned_words.sort(key=lambda item: item[0])
    words = [word for _, word in positioned_words]
    text = " ".join(words)
    text = re.sub(r"\s+([,.;:?!])", r"\1", text)
    text = re.sub(r"\(\s+", "(", text)
    text = re.sub(r"\s+\)", ")", text)
    return text.strip()


def fetch_abstract_from_europe_pmc(title: str, journal: str, doi: str, mailto: str | None) -> str:
    queries: list[str] = []
    if doi:
        queries.append(f'DOI:"{doi}"')
    title_query = build_europe_pmc_title_query(title, journal)
    if title_query:
        queries.append(title_query)

    headers = {}
    if mailto:
        headers["User-Agent"] = f"JournalFetcher/1.0 (mailto:{mailto})"

    for query in queries:
        params = {
            "query": query,
            "resultType": "core",
            "format": "json",
            "pageSize": "1",
        }
        request = urllib.request.Request(
            f"{EUROPE_PMC_SEARCH_URL}?{urllib.parse.urlencode(params)}",
            headers=headers,
        )
        try:
            with urllib.request.urlopen(request, timeout=20) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError):
            continue

        results = payload.get("resultList", {}).get("result", [])
        if not results:
            continue
        abstract = clean_abstract(results[0].get("abstractText"))
        if abstract:
            return abstract
    return ""


def build_europe_pmc_title_query(title: str, journal: str) -> str:
    clean_title = normalize_query_text(title)
    clean_journal = normalize_query_text(journal)
    if not clean_title:
        return ""
    if clean_journal:
        return f'TITLE:"{clean_title}" AND JOURNAL:"{clean_journal}"'
    return f'TITLE:"{clean_title}"'


def normalize_query_text(value: str) -> str:
    value = re.sub(r"\s+", " ", value or "").strip()
    return value.replace('"', "")


def journal_matches(user_input: str, candidate: str) -> bool:
    if not candidate:
        return False
    left = normalize_text(user_input)
    right = normalize_text(candidate)
    if not left or not right:
        return False
    if left in right or right in left:
        return True
    return len(meaningful_tokens(user_input) & meaningful_tokens(candidate)) >= 1


def normalize_text(value: str) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", value.lower())


def meaningful_tokens(value: str) -> set[str]:
    return {token for token in re.findall(r"[A-Za-z0-9\u4e00-\u9fff]+", value.lower()) if len(token) > 1}


def first_text(values: Iterable[str]) -> str:
    for value in values:
        if value:
            return str(value).strip()
    return ""


def extract_date(item: dict) -> str:
    for key in ("published-print", "published-online", "issued"):
        parts = item.get(key, {}).get("date-parts", [])
        if parts and parts[0]:
            values = [str(part) for part in parts[0]]
            while len(values) < 3:
                values.append("01")
            return f"{values[0]}-{values[1].zfill(2)}-{values[2].zfill(2)}"
    return ""


def clean_abstract(raw_abstract: str | None) -> str:
    if not raw_abstract:
        return ""
    text = re.sub(r"<[^>]+>", " ", raw_abstract)
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def resolve_output_path(output: str | None) -> Path:
    if output:
        path = Path(output)
    else:
        path = Path(__file__).resolve().parent / f"journal_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def save_latest_copy(output_path: Path) -> Path:
    latest_path = output_path.parent / "journal_results_latest.xlsx"
    shutil.copyfile(output_path, latest_path)
    return latest_path


def write_excel(rows: list[ArticleRow], output_path: Path) -> None:
    table = [["标题", "期刊名", "日期", "主题", "摘要"]]
    table.extend([[r.title, r.journal, r.date, r.topic, r.abstract] for r in rows])
    strings, index = build_shared_strings(table)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", build_content_types_xml())
        archive.writestr("_rels/.rels", build_rels_xml())
        archive.writestr("xl/workbook.xml", build_workbook_xml())
        archive.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels_xml())
        archive.writestr("xl/styles.xml", build_styles_xml())
        archive.writestr("xl/sharedStrings.xml", build_shared_strings_xml(strings))
        archive.writestr("xl/worksheets/sheet1.xml", build_sheet_xml(table, index))


def build_shared_strings(rows: list[list[str]]) -> tuple[list[str], dict[str, int]]:
    values: list[str] = []
    index: dict[str, int] = {}
    for row in rows:
        for cell in row:
            text = cell or ""
            if text not in index:
                index[text] = len(values)
                values.append(text)
    return values, index


def build_sheet_xml(rows: list[list[str]], shared_index: dict[str, int]) -> str:
    row_xml: list[str] = []
    for row_num, row in enumerate(rows, start=1):
        cells: list[str] = []
        for col_num, value in enumerate(row, start=1):
            cell_ref = f"{column_letter(col_num)}{row_num}"
            style = ' s="1"' if row_num == 1 else ""
            cells.append(f'<c r="{cell_ref}" t="s"{style}><v>{shared_index[value or ""]}</v></c>')
        row_xml.append(f'<row r="{row_num}">{"".join(cells)}</row>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="42" customWidth="1"/><col min="2" max="2" width="28" customWidth="1"/><col min="3" max="3" width="14" customWidth="1"/><col min="4" max="4" width="20" customWidth="1"/><col min="5" max="5" width="90" customWidth="1"/></cols><sheetData>' + "".join(row_xml) + "</sheetData></worksheet>"


def build_shared_strings_xml(strings: list[str]) -> str:
    items = "".join(shared_string_node(value) for value in strings)
    count = len(strings)
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{count}" uniqueCount="{count}">{items}</sst>'


def shared_string_node(value: str) -> str:
    text = escape(value or "")
    if text[:1].isspace() or text[-1:].isspace():
        return f'<si><t xml:space="preserve">{text}</t></si>'
    return f"<si><t>{text}</t></si>"


def build_workbook_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="文献结果" sheetId="1" r:id="rId1"/></sheets></workbook>'


def build_rels_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'


def build_workbook_rels_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'


def build_styles_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'


def build_content_types_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'


def column_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def print_preview(rows: list[ArticleRow]) -> None:
    if not rows:
        print("没有检索到符合条件的结果。")
        return
    print(f"共检索到 {len(rows)} 条结果，下面预览前 5 条：")
    for i, row in enumerate(rows[:5], start=1):
        print(f"[{i}] {row.title}")
        print(f"    期刊: {row.journal}")
        print(f"    日期: {row.date}")
        print(f"    主题: {row.topic}")
        print(f"    摘要: {truncate(row.abstract, 180)}")


def truncate(text: str, max_length: int) -> str:
    return text if len(text) <= max_length else text[: max_length - 3].rstrip() + "..."


def run_cli(args: argparse.Namespace) -> int:
    journal = prompt_if_missing(args.journal, "请输入期刊名称")
    topic = prompt_if_missing(args.topic, "请输入文献主题（可直接回车跳过）", required=False)
    from_date = validate_date(prompt_if_missing(args.from_date, "请输入开始日期（YYYY-MM-DD）", required=False), "开始日期")
    until_date = validate_date(prompt_if_missing(args.until_date, "请输入结束日期（YYYY-MM-DD）", required=False), "结束日期")
    try:
        rows = fetch_articles(journal or "", topic, from_date, until_date, args.max_results, args.mailto)
    except urllib.error.HTTPError as exc:
        print(f"请求失败：HTTP {exc.code}", file=sys.stderr)
        return 1
    except urllib.error.URLError as exc:
        print(f"网络请求异常：{exc.reason}", file=sys.stderr)
        return 1
    output_path = resolve_output_path(args.output)
    write_excel(rows, output_path)
    latest_path = save_latest_copy(output_path)
    print_preview(rows)
    print(f"\n结果已保存到：{output_path}")
    print(f"固定文件也已更新：{latest_path}")
    return 0


def render_page(form: dict[str, str] | None = None, output_path: str = "", count: int = 0, error_message: str = "") -> str:
    data = form or {}
    ok_style = "block" if output_path else "none"
    err_style = "block" if error_message else "none"
    return f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Journal Fetcher</title><style>body{{font-family:Segoe UI,Microsoft YaHei,sans-serif;background:#f5efe6;color:#21313c;margin:0}}main{{max-width:1080px;margin:0 auto;padding:24px;display:grid;grid-template-columns:1.1fr .9fr;gap:18px}}section{{background:#fffaf2;border:1px solid #d9cab7;border-radius:16px;padding:20px;box-shadow:0 12px 28px rgba(0,0,0,.06)}}h1{{margin:0 0 10px;font-size:30px}}p{{line-height:1.6;color:#596574}}label{{display:block;margin:12px 0 6px;font-weight:600}}input,textarea{{width:100%;padding:10px 12px;border:1px solid #d9cab7;border-radius:10px;background:#fffdf8;box-sizing:border-box}}textarea{{min-height:320px;white-space:pre-wrap}}button{{margin-top:16px;background:#17624a;color:#fff;border:none;border-radius:999px;padding:11px 18px;cursor:pointer}}.msg{{margin-top:14px;padding:12px;border-radius:10px}}.ok{{background:#e7f5ef;display:{ok_style}}}.err{{background:#fdeaea;color:#7b2020;display:{err_style}}}.mono{{font-family:Consolas,monospace;word-break:break-all}}@media(max-width:900px){{main{{grid-template-columns:1fr}}}}</style></head><body><main><section><h1>Journal Fetcher</h1><p>输入期刊、时间范围和可选主题，导出 Excel。主题留空时，会抓取该时间段内该期刊的所有文章。</p><form method="post" action="/search"><label>期刊名称</label><input name="journal" value="{html.escape(data.get('journal', ''))}" required><label>文献主题（可留空）</label><input name="topic" value="{html.escape(data.get('topic', ''))}"><label>开始日期</label><input name="from_date" placeholder="YYYY-MM-DD" value="{html.escape(data.get('from_date', ''))}"><label>结束日期</label><input name="until_date" placeholder="YYYY-MM-DD" value="{html.escape(data.get('until_date', ''))}"><label>最多结果数（0 表示全部）</label><input name="max_results" type="number" min="0" max="5000" value="{html.escape(data.get('max_results', '0'))}"><label>联系邮箱（可选）</label><input name="mailto" value="{html.escape(data.get('mailto', ''))}"><button type="submit">开始检索并导出 Excel</button></form><div class="msg ok"><strong>导出完成</strong><p>共整理 {count} 条结果。</p><p>文件位置：<span class="mono">{html.escape(output_path)}</span></p></div><div class="msg err"><strong>处理失败</strong><p>{html.escape(error_message)}</p></div></section><section><h1 style="font-size:24px">使用说明</h1><textarea readonly>{html.escape(README_TEXT)}</textarea></section></main></body></html>"""


def make_handler() -> type[BaseHTTPRequestHandler]:
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:
            self.respond(render_page())

        def do_POST(self) -> None:
            if self.path != "/search":
                self.send_error(404)
                return
            length = int(self.headers.get("Content-Length", "0"))
            form = {k: v[0] for k, v in urllib.parse.parse_qs(self.rfile.read(length).decode("utf-8")).items()}
            try:
                journal = (form.get("journal") or "").strip()
                if not journal:
                    raise ValueError("期刊名称不能为空。")
                topic = (form.get("topic") or "").strip() or None
                from_date = validate_date((form.get("from_date") or "").strip() or None, "开始日期")
                until_date = validate_date((form.get("until_date") or "").strip() or None, "结束日期")
                max_results = int((form.get("max_results") or "20").strip())
                if max_results < 0:
                    raise ValueError("最多结果数不能小于 0。")
                rows = fetch_articles(journal, topic, from_date, until_date, max_results, (form.get("mailto") or "").strip() or None)
                output_path = resolve_output_path(None)
                write_excel(rows, output_path)
                latest_path = save_latest_copy(output_path)
                self.respond(render_page(form, f"{output_path} | 固定文件: {latest_path}", len(rows), ""))
            except ValueError as exc:
                self.respond(render_page(form, "", 0, str(exc)), 400)
            except urllib.error.HTTPError as exc:
                self.respond(render_page(form, "", 0, f"请求失败：HTTP {exc.code}"), 502)
            except urllib.error.URLError as exc:
                self.respond(render_page(form, "", 0, f"网络请求异常：{exc.reason}"), 502)
            except Exception as exc:
                self.respond(render_page(form, "", 0, f"未预期错误：{exc}"), 500)

        def log_message(self, format: str, *args: object) -> None:
            return

        def respond(self, content: str, status: int = 200) -> None:
            data = content.encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

    return Handler


def run_web(port: int, auto_open_browser: bool) -> int:
    server = ThreadingHTTPServer(("127.0.0.1", port), make_handler())
    url = f"http://127.0.0.1:{port}"
    print(f"网页端已启动：{url}")
    if auto_open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止网页端。")
    finally:
        server.server_close()
    return 0


def main() -> int:
    args = parse_args()
    return run_web(args.port, not args.no_browser) if args.web else run_cli(args)


if __name__ == "__main__":
    raise SystemExit(main())
