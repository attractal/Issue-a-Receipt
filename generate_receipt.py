#!/usr/bin/env python3

# モジュールのインポート
# 環境に合わせて、ないものは pip でインストールしてください。
from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

# たぶん lxml は pip install lxml で入ります。
from lxml import etree


# Word の DOCX は ZIP ファイルの中に XML が入っている形式です。
# ここでは Word 本文の XML に出てくる「w:t」というテキスト要素だけを差し替えます。
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

# 何も指定しない場合は、この Python ファイルと同じフォルダにある
# receipt_template_source.docx をテンプレートとして使います。
TEMPLATE = Path(__file__).with_name("receipt_template_source.docx")


@dataclass
class Receipt:
    """領収書 1 枚分の入力データをまとめて持つクラスです。"""

    # dataclass を使うと、コンストラクタや __repr__ などを自動生成してくれます。
    recipient: str  # 宛名
    amount: int  # 税込金額（整数、円単位）
    issued_on: date  # 発行日
    description: str  # 但し書き
    staff: str = "ScienceTechno"  # 担当者名（省略時は「ScienceTechno」）

    # 税込金額から税抜金額と消費税等を計算するプロパティを用意します。
    @property
    def net_amount(self) -> int:
        # 税抜金額 = 税込金額 - 消費税等
        return self.amount - self.tax_amount

    # 消費税等は、税込金額の 10/110 で計算します。
    # TODO: 消費税率が変わった場合は、ここ (110) を修正してください。
    @property
    def tax_amount(self) -> int:
        # 元テンプレートの計算に合わせ、消費税は端数切り捨てにしています。
        # 例: 10,000 円 -> 消費税 909 円、税抜 9,091 円
        return (
            self.amount * 10 // 110
        )  # <- TODO: 消費税率が変わった場合は、ここ (10/110) を修正してください。


def format_date(value: date) -> str:
    """date 型の日付を、領収書に表示する「2004年6月27日」の形にします。"""

    return f"{value.year}年{value.month}月{value.day}日"


def parse_date(value: str | None) -> date:
    """文字列の日付を date 型に変換します。未指定なら今日の日付を使います。"""

    if not value:
        return date.today()

    # 利用者が入力しやすいように、いくつかの日付形式を受け付けます。
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Unsupported date format: {value}. Use YYYY-MM-DD.")


def parse_amount(value: str | int) -> int:
    """金額を整数に変換します。「10,000」や「¥10000」も受け付けます。"""

    if isinstance(value, int):
        return value
    normalized = value.replace(",", "").replace("¥", "").replace("\\", "").strip()
    amount = int(normalized)
    if amount < 0:
        raise ValueError("amount must be zero or greater")
    return amount


def currency(value: int, leading_space: bool = True) -> str:
    """金額にカンマを付けます。例: 9091 -> 9,091"""

    text = f"{value:,}"
    return f" {text}" if leading_space else text


def text_nodes(element: etree._Element) -> list[etree._Element]:
    """XML の中から、実際の文字が入っている w:t 要素だけを取り出します。"""

    return element.xpath(".//w:t", namespaces=NS)


def set_texts(nodes: list[etree._Element], values: list[str]) -> None:
    """既存の Word テキスト要素に、新しい文字列を順番に入れます。"""

    if not nodes:
        raise ValueError("template block does not contain editable text")

    # テンプレートより差し替え後の文字の分割数が多い場合に備えて、
    # 最後のテキスト要素をコピーして不足分を作ります。
    while len(nodes) < len(values):
        clone = etree.Element(nodes[-1].tag, nsmap=nodes[-1].nsmap)
        nodes[-1].addnext(clone)
        nodes.append(clone)
    for node, value in zip(nodes, values):
        node.text = value
    for node in nodes[len(values) :]:
        node.text = ""


def fill_receipt_block(blocks: list[etree._Element], receipt: Receipt) -> None:
    """領収書 1 枚分の XML ブロックに、入力データを差し込みます。"""

    for block in blocks:
        # 1 つの段落や表の中の文字をつなげて、どの項目か判定します。
        joined = "".join(node.text or "" for node in text_nodes(block))
        nodes = text_nodes(block)

        if joined.startswith("発行日"):
            set_texts(
                nodes, ["発行日", f"\u3000{format_date(receipt.issued_on)}\u3000"]
            )
        elif "御中" in joined:
            set_texts(nodes, [f"\u3000\u3000{receipt.recipient}", "\u3000\u3000御中"])
        elif joined.startswith("¥"):
            set_texts(nodes, ["¥    ", f"{receipt.amount:,}", " -（税込）"])
        elif joined.startswith("但"):
            set_texts(nodes, ["但\u3000", receipt.description, "として"])
        elif joined.startswith("内訳"):
            set_texts(
                nodes,
                [
                    "内訳",
                    "税抜金額",
                    "\\",
                    currency(receipt.net_amount),
                    "消費税等",
                    "\\",
                    currency(receipt.tax_amount),
                ],
            )
        elif joined.startswith("担当:"):
            set_texts(nodes, [f"担当: {receipt.staff}\u3000\u3000\u3000"])


def fill_docx(template: Path, output: Path, receipts: list[Receipt]) -> None:
    """テンプレート DOCX を読み込み、領収書データを差し込んだ DOCX を保存します。"""

    if not template.exists():
        raise FileNotFoundError(f"Template not found: {template}")
    if not receipts:
        raise ValueError("at least one receipt is required")

    with ZipFile(template) as zin:
        # DOCX の本文は word/document.xml に入っています。
        document_xml = zin.read("word/document.xml")
        root = etree.fromstring(document_xml)
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("template does not contain a Word body")

        blocks = [child for child in body if child.tag.endswith(("}p", "}tbl"))]

        # テンプレートには「領収書」という見出しが上下 2 つあります。
        # この見出し位置を目印に、上段と下段の範囲を分けます。
        title_indexes = [
            i
            for i, block in enumerate(blocks)
            if "".join(node.text or "" for node in text_nodes(block)) == "領収書"
        ]
        if not title_indexes:
            raise ValueError("could not find receipt sections in template")

        section_ranges = []
        for pos, start in enumerate(title_indexes):
            end = (
                title_indexes[pos + 1] if pos + 1 < len(title_indexes) else len(blocks)
            )
            section_ranges.append((start, end))

        for idx, (start, end) in enumerate(section_ranges):
            # 入力が 1 件だけの場合は、上下とも同じ内容にします。
            # CSV で 2 件渡した場合は、1 件目が上段、2 件目が下段になります。
            receipt = receipts[min(idx, len(receipts) - 1)]
            fill_receipt_block(blocks[start:end], receipt)

        rendered_xml = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        output.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(output, "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                # document.xml だけ差し替え、それ以外の画像・スタイル・余白などは
                # 元テンプレートのままコピーします。
                data = (
                    rendered_xml
                    if item.filename == "word/document.xml"
                    else zin.read(item.filename)
                )
                zout.writestr(item, data)


def load_csv(path: Path) -> list[Receipt]:
    """CSV ファイルを読み込み、Receipt のリストに変換します。"""

    receipts: list[Receipt] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        # DictReader を使うと、列名で値を取り出せます。
        for row in csv.DictReader(f):
            receipts.append(
                Receipt(
                    recipient=row["recipient"],
                    amount=parse_amount(row["amount"]),
                    issued_on=parse_date(row.get("date")),
                    description=row["description"],
                    staff=row.get("staff") or "ScienceTechno",
                )
            )
    return receipts


def receipt_from_args(args: argparse.Namespace) -> Receipt:
    """コマンドライン引数から、領収書 1 件分の Receipt を作ります。"""

    missing = [
        name
        for name in ("recipient", "amount", "description")
        if not getattr(args, name)
    ]
    if missing:
        raise ValueError(
            f"missing required option(s): {', '.join('--' + name for name in missing)}"
        )
    return Receipt(
        recipient=args.recipient,
        amount=parse_amount(args.amount),
        issued_on=parse_date(args.date),
        description=args.description,
        staff=args.staff,
    )


def main() -> None:
    # argparse は「--recipient」などのコマンドライン引数を読み取るための標準ライブラリです。
    parser = argparse.ArgumentParser(
        description="Issue a receipt DOCX from the receipt template."
    )
    parser.add_argument("--recipient", help="宛名。")
    parser.add_argument("--amount", help="税込金額。例: 10000")
    parser.add_argument("--description", help="但し書き。")
    parser.add_argument(
        "--staff", default="ScienceTechno", help="担当名。省略時は ScienceTechno"
    )
    parser.add_argument("--date", help="発行日。YYYY-MM-DD。省略時は今日")
    parser.add_argument(
        "--csv",
        type=Path,
        help="CSV input with recipient,amount,description,date,staff columns",
    )
    parser.add_argument(
        "--template", type=Path, default=TEMPLATE, help="template DOCX path"
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("issued_receipt.docx"),
        help="output DOCX path",
    )
    parser.add_argument(
        "--keep-one-per-file",
        action="store_true",
        help="with --csv, create one DOCX per row instead of filling one page with up to two receipts",
    )
    args = parser.parse_args()

    if args.csv:
        # CSV が指定された場合は、CSV の各行を領収書データとして読み込みます。
        receipts = load_csv(args.csv)
        if args.keep_one_per_file:
            # --keep-one-per-file の場合は、CSV の各行を別々の DOCX にします。
            stem = args.output.stem
            suffix = args.output.suffix or ".docx"
            for i, receipt in enumerate(receipts, start=1):
                fill_docx(
                    args.template,
                    args.output.with_name(f"{stem}_{i:03d}{suffix}"),
                    [receipt],
                )
        else:
            # 通常は、CSV の 1 行目を上段、2 行目を下段に差し込みます。
            fill_docx(args.template, args.output, receipts[:2])
    else:
        # CSV が無い場合は、コマンドライン引数から 1 件分だけ作ります。
        fill_docx(args.template, args.output, [receipt_from_args(args)])


if __name__ == "__main__":
    main()
