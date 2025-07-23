import argparse
import re
import os
import html
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from difflib import ndiff

def parse_args():
    parser = argparse.ArgumentParser(description="Generate HTML diff report from Excel files.")
    parser.add_argument("--original", required=True, help="Path to original file or directory")
    parser.add_argument("--modified", required=True, help="Path to modified file or directory")
    parser.add_argument("--wspattern", required=True, help="Regex pattern to match worksheet names")
    parser.add_argument("--source", required=True, help="Source column (e.g., B)")
    parser.add_argument("--target", required=True, help="Target column (e.g., C)")
    parser.add_argument("--row-offset", type=int, default=0, help="Rows to skip (e.g., header rows)")
    parser.add_argument("--word", action="store_true", help="Diff by word (default)")
    parser.add_argument("--char", action="store_true", help="Diff by character")
    parser.add_argument("--output", help="Output HTML file (default: diff_report.html)")
    parser.add_argument("--dir", action="store_true", help="Directory mode - compare folders")
    return parser.parse_args()

def get_diff_html(orig, mod, by_word=True):
    orig = orig or ""
    mod = mod or ""
    tokenize = (lambda x: x.split()) if by_word else list
    diff = ndiff(tokenize(orig), tokenize(mod))
    result = []
    for d in diff:
        op, text = d[0], html.escape(d[2:])
        if op == ' ':
            result.append(text)
        elif op == '+':
            result.append(f'<span style="color:green;">{text}</span>')
        elif op == '-':
            result.append(f'<span style="color:red;text-decoration:line-through;">{text}</span>')
    return (' ' if by_word else '').join(result)

def process_sheet(ws_orig, ws_mod, source_col, target_col, row_offset, by_word):
    max_row = max(ws_orig.max_row, ws_mod.max_row)
    rows = []
    for r in range(row_offset + 1, max_row + 1):
        source = ws_orig.cell(row=r, column=source_col).value or ""
        orig_val = ws_orig.cell(r, column=target_col).value or ""
        mod_val  = ws_mod.cell(r, column=target_col).value or ""
        if orig_val != mod_val:
            diff_view = get_diff_html(orig_val, mod_val, by_word)
            rows.append((r, source, orig_val, mod_val, diff_view))
    return rows

def process_sheet_as_deletion(ws, source_col, target_col, row_offset, by_word):
    rows = []
    for r in range(row_offset + 1, ws.max_row + 1):
        source = ws.cell(r, column=source_col).value or ""
        orig_val = ws.cell(r, column=target_col).value or ""
        mod_val = ""
        if orig_val:
            diff_view = get_diff_html(orig_val, mod_val, by_word)
            rows.append((r, source, orig_val, mod_val, diff_view))
    return rows

def process_sheet_as_insertion(ws, source_col, target_col, row_offset, by_word):
    rows = []
    for r in range(row_offset + 1, ws.max_row + 1):
        source = ws.cell(r, column=source_col).value or ""
        orig_val = ""
        mod_val = ws.cell(r, column=target_col).value or ""
        if mod_val:
            diff_view = get_diff_html(orig_val, mod_val, by_word)
            rows.append((r, source, orig_val, mod_val, diff_view))
    return rows

def compare_workbook_pair(file_label, original_path, modified_path, wspattern, source_col, target_col, row_offset, by_word):
    wb_o = load_workbook(original_path, data_only=True)
    wb_m = load_workbook(modified_path, data_only=True)
    sheet_results = []

    for name in wb_o.sheetnames:
        if not wspattern.match(name) or name not in wb_m.sheetnames:
            continue
        rows = process_sheet(wb_o[name], wb_m[name], source_col, target_col, row_offset, by_word)
        if rows:
            sheet_results.append((name, rows))
    return (file_label, sheet_results) if sheet_results else None

def write_html_report(filename, all_results, use_dir_mode):
    with open(filename, "w", encoding="utf-8") as f:
        f.write("<!DOCTYPE html><html><head><meta charset='utf-8'>\n")
        f.write("<title>Excel Diff Report</title>\n")
        f.write("""
        <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            table { border-collapse: collapse; margin-bottom: 2em; width: 100%; }
            th, td { border: 1px solid #ccc; padding: 8px; vertical-align: top; }
            th { background: #f0f0f0; }
            a { color: #0077cc; }
            .toc ul { list-style-type: none; padding-left: 1em; }
            .toc li { margin-bottom: 4px; }
            .toc .file { font-weight: bold; margin-top: 1em; }
            pre { white-space: pre-wrap; }
        </style>
        </head><body>
        <h1>Excel Diff Report</h1>
        <div class="toc"><h2>Table of Contents</h2>
        <ul>
        """)
        for file_label, sheets in all_results:
            anchor = f"file--{html.escape(file_label)}"
            f.write(f"<li class='file'><a href='#{anchor}'>{html.escape(file_label)}</a></li>\n<ul>\n")
            for sheet_name, _ in sheets:
                sheet_anchor = f"{anchor}--{sheet_name.replace(' ', '_')}"
                f.write(f"<li><a href='#{sheet_anchor}'>{html.escape(sheet_name)}</a></li>\n")
            f.write("</ul>\n")
        f.write("</ul></div>\n")

        for file_label, sheets in all_results:
            anchor = f"file--{html.escape(file_label)}"
            f.write(f"<h2 id='{anchor}'>{html.escape(file_label)}</h2>\n")
            for sheet_name, rows in sheets:
                sheet_anchor = f"{anchor}--{sheet_name.replace(' ', '_')}"
                f.write(f"<h3 id='{sheet_anchor}'>{html.escape(sheet_name)}</h3>\n")
                f.write("<table>\n")
                f.write("<tr><th>Line</th><th>Source</th><th>Original Target</th><th>Modified Target</th><th>Diff</th></tr>\n")
                for row in rows:
                    line, source, orig, mod, diff = row
                    f.write(f"<tr><td>{line}</td><td>{html.escape(str(source))}</td>"
                            f"<td>{html.escape(str(orig))}</td><td>{html.escape(str(mod))}</td>"
                            f"<td><pre>{diff}</pre></td></tr>\n")
                f.write("</table>\n")
        f.write("</body></html>")

def main():
    args = parse_args()
    by_word = args.word or not args.char
    wspattern = re.compile(args.wspattern)
    source_col = column_index_from_string(args.source)
    target_col = column_index_from_string(args.target)
    row_offset = args.row_offset
    out_file = args.output or "diff_report.html"
    results = []

    if args.dir:
        orig_dir = args.original
        mod_dir = args.modified

        orig_files = {
            f for f in os.listdir(orig_dir)
            if f.lower().endswith((".xls", ".xlsx")) and os.path.isfile(os.path.join(orig_dir, f))
        }
        mod_files = {
            f for f in os.listdir(mod_dir)
            if f.lower().endswith((".xls", ".xlsx")) and os.path.isfile(os.path.join(mod_dir, f))
        }

        all_filenames = sorted(orig_files | mod_files)

        for filename in all_filenames:
            orig_path = os.path.join(orig_dir, filename)
            mod_path = os.path.join(mod_dir, filename)

            has_orig = filename in orig_files
            has_mod = filename in mod_files

            if has_orig and has_mod:
                result = compare_workbook_pair(filename, orig_path, mod_path,
                                               wspattern, source_col, target_col, row_offset, by_word)
                if result:
                    results.append(result)
            elif has_orig:
                wb = load_workbook(orig_path, data_only=True)
                sheet_results = []
                for sheet in wb.sheetnames:
                    if wspattern.match(sheet):
                        rows = process_sheet_as_deletion(wb[sheet], source_col, target_col, row_offset, by_word)
                        if rows:
                            sheet_results.append((sheet, rows))
                if sheet_results:
                    results.append((f"{filename} (deleted)", sheet_results))
            elif has_mod:
                wb = load_workbook(mod_path, data_only=True)
                sheet_results = []
                for sheet in wb.sheetnames:
                    if wspattern.match(sheet):
                        rows = process_sheet_as_insertion(wb[sheet], source_col, target_col, row_offset, by_word)
                        if rows:
                            sheet_results.append((sheet, rows))
                if sheet_results:
                    results.append((f"{filename} (added)", sheet_results))
    else:
        label = os.path.basename(args.original)
        result = compare_workbook_pair(label, args.original, args.modified,
                                       wspattern, source_col, target_col, row_offset, by_word)
        if result:
            results.append(result)

    if results:
        write_html_report(out_file, results, use_dir_mode=args.dir)
        print(f"✅ Report saved: {out_file}")
    else:
        print("⚠️ No differences found or matching files/sheets.")

if __name__ == "__main__":
    main()
