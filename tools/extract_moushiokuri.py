#!/usr/bin/env python3
"""介護システムから落としたExcelから、申し送り欄だけを取り出す。

利用者名は自動でマスクする（山田花子 → 利用者A）。
取り出した結果は work/ に置き、そこから教材を作る。

使い方:
    python3 tools/extract_moushiokuri.py 記録.xlsx --list        # 何の列があるか見る
    python3 tools/extract_moushiokuri.py 記録.xlsx               # 申し送りを取り出す
    python3 tools/extract_moushiokuri.py 記録.xlsx --since 2026-08-12
"""
import argparse
import datetime as dt
import os
import re
import sys
import unicodedata

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl が必要です:  pip install openpyxl")

# 列の見つけ方。上にあるものほど優先して使う。
HEADER_HINTS = {
    "text": ["申し送り", "申送り", "申送", "引き継ぎ", "引継ぎ", "引継",
             "特記事項", "経過記録", "記録内容", "本文", "内容", "コメント", "備考"],
    "date": ["記録日", "日付", "日時", "年月日", "実施日"],
    "name": ["利用者名", "入居者名", "ご利用者", "利用者", "氏名", "お名前", "名前"],
    "staff": ["記録者", "担当者", "職員", "記入者", "スタッフ"],
}
MAX_HEADER_SCAN_ROWS = 15   # ヘッダー行はふつう最初の数行にある
MASK_LABELS = [chr(ord("A") + i) for i in range(26)]


def norm(v):
    """セルの値を、比較しやすい文字列に整える。"""
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return unicodedata.normalize("NFKC", str(v)).strip()


def find_header_row(rows):
    """申し送りらしき列名を含む行を、ヘッダー行とみなす。"""
    best, best_score = None, 0
    for i, row in enumerate(rows[:MAX_HEADER_SCAN_ROWS]):
        cells = [norm(c) for c in row]
        score = sum(1 for kind in HEADER_HINTS
                    for hint in HEADER_HINTS[kind]
                    if any(hint in c for c in cells if c))
        if score > best_score:
            best, best_score = i, score
    return best


def match_column(headers, kind):
    """ヘッダー名から、その種類の列番号を返す。見つからなければ None。"""
    for hint in HEADER_HINTS[kind]:
        for idx, h in enumerate(headers):
            if hint and hint in h:
                return idx
    return None


def load_sheet(path, sheet_name=None):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    names = [sheet_name] if sheet_name else wb.sheetnames
    for name in names:
        if name not in wb.sheetnames:
            sys.exit(f"シート「{name}」がありません。あるのは: {', '.join(wb.sheetnames)}")
        rows = [list(r) for r in wb[name].iter_rows(values_only=True)]
        if not rows:
            continue
        hi = find_header_row(rows)
        if hi is None:
            continue
        headers = [norm(c) for c in rows[hi]]
        if match_column(headers, "text") is not None:
            return name, headers, rows[hi + 1:], wb.sheetnames
    return None, None, None, wb.sheetnames


def build_mask(user_names, staff_names):
    """氏名 → 利用者A/職員A… の対応表を作る。長い名前から先に置換する。"""
    mapping = {}
    for names, label in [(user_names, "利用者"), (staff_names, "職員")]:
        seen = []
        for n in names:
            n = n.strip()
            if len(n) >= 2 and n not in mapping:
                mapping[n] = f"{label}{MASK_LABELS[len(seen) % 26]}"
                seen.append(n)
    # 「山田花子」を先に消してから「山田」を消す
    order = sorted(mapping, key=len, reverse=True)
    return mapping, order


def apply_mask(text, mapping, order):
    for name in order:
        text = text.replace(name, mapping[name])
        # 姓だけ・名だけで書かれている場合も拾う
        for part in re.split(r"[\s　]+", name):
            if len(part) >= 2:
                text = text.replace(part, mapping[name])
    return text


def leftover_names(text):
    """マスクし残した「〇〇様」「〇〇さん」を警告として拾う。"""
    found = re.findall(r"([一-龥ぁ-んァ-ヶ]{2,5})(?:様|さん)", text)
    return [f for f in found if not f.startswith(("利用者", "職員"))]


def main():
    ap = argparse.ArgumentParser(description="Excelから申し送り欄を取り出す")
    ap.add_argument("excel", help="介護システムから落としたExcelファイル")
    ap.add_argument("--sheet", help="シート名（省略時は自動で探す）")
    ap.add_argument("--list", action="store_true", help="列名だけ表示して終わる")
    ap.add_argument("--since", help="この日以降だけ取り出す (YYYY-MM-DD)")
    ap.add_argument("--out", help="出力先（省略時は work/<日付>-申し送り.md）")
    ap.add_argument("--no-mask", action="store_true",
                    help="名前をマスクしない（施設内での確認用。送信物には使わないこと）")
    args = ap.parse_args()

    if not os.path.exists(args.excel):
        sys.exit(f"ファイルが見つかりません: {args.excel}")

    sheet, headers, rows, all_sheets = load_sheet(args.excel, args.sheet)
    if sheet is None:
        sys.exit("申し送りらしき列が見つかりませんでした。\n"
                 f"シート一覧: {', '.join(all_sheets)}\n"
                 "--list で列名を確認し、HEADER_HINTS に列名を追加してください。")

    ci = {k: match_column(headers, k) for k in HEADER_HINTS}

    if args.list:
        print(f"シート: {sheet}")
        print(f"列: {', '.join(h for h in headers if h)}")
        print("\n自動で選んだ列:")
        for k, label in [("text", "申し送り"), ("date", "日付"), ("name", "利用者名"), ("staff", "記録者")]:
            print(f"  {label}: {headers[ci[k]] if ci[k] is not None else '(なし)'}")
        return

    def column_values(kind):
        idx = ci[kind]
        if idx is None:
            return []
        return [norm(r[idx]) for r in rows if idx < len(r)]

    mapping, order = build_mask(column_values("name"), column_values("staff"))

    entries, warnings = [], set()
    for r in rows:
        text = norm(r[ci["text"]]) if ci["text"] < len(r) else ""
        if not text:
            continue
        date = norm(r[ci["date"]]) if ci["date"] is not None and ci["date"] < len(r) else ""
        if args.since and date and date < args.since:
            continue
        if not args.no_mask:
            text = apply_mask(text, mapping, order)
            warnings.update(leftover_names(text))
        entries.append((date, text))

    if not entries:
        sys.exit("申し送りの中身が1件も取れませんでした。--list で列名を確認してください。")

    out = args.out or os.path.join(
        "work", f"{dt.date.today():%Y-%m-%d}-申し送り.md")
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(f"# 申し送り抽出（{len(entries)}件）\n\n")
        f.write(f"- 元ファイル: {os.path.basename(args.excel)} / シート: {sheet}\n")
        f.write(f"- 抽出列: {headers[ci['text']]}\n")
        f.write(f"- 名前マスク: {'なし' if args.no_mask else f'あり（{len(mapping)}名）'}\n\n")
        f.write("> このファイルは教材を作るための作業用です。そのまま実習生に送らないこと。\n\n")
        for date, text in entries:
            f.write(f"- **{date or '日付なし'}**　{text}\n")

    print(f"{len(entries)}件を書き出しました → {out}")
    if warnings:
        print("\n⚠ マスクし残しの可能性がある名前があります。教材にする前に確認してください:")
        for w in sorted(warnings):
            print(f"   - {w}様／{w}さん")


if __name__ == "__main__":
    main()
