#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JAL予約メール → 予約経路チェック図 HTML生成ツール

役割:
  1. JALMailEngine.py でJAL予約メールを取得・解析する
  2. 搭乗日/出発時刻順にフライトを並べる
  3. jal_route_map.tmpl の __TIMELINE_FLIGHTS_JSON__ へJSONを差し込む
  4. 時系列で移動経路を確認できるHTMLを出力する

注意:
  - SVG、CSS、UI、JavaScriptの固定部分は jal_route_map.tmpl 側に集約
  - 本体Pythonは「データ生成」と「テンプレート差し込み」に限定
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, List

import JALMailEngine


# ============================================================
# 設定
# ============================================================

IMAP_HOST = "imap.example.com"
IMAP_PORT = 993
IMAP_USER = "your_mail@example.com"
IMAP_PASSWORD = "your_password"
IMAP_FOLDER = "INBOX"

# 対象とするJAL購入通知メール件名
MAIL_SUBJECT_REGEX = r"^〔JAL国内線〕購入内容のお知らせ"

BASE_NETWORK_TEMPLATE_PATH = "jal_route_map.tmpl"
LOGO_PLACEHOLDER = "__LOGO_IMAGE_ELEMENT__"
OUTPUT_HTML = "jal_route_timeline.html"


# ============================================================
# 空港名変換
# ============================================================

SVG_AIRPORT_ALIASES = {
    "東京": "羽田",
    "東京(羽田)": "羽田",
    "羽田": "羽田",
    "東京(成田)": "成田",
    "成田": "成田",

    "大阪": "伊丹",
    "大阪(伊丹)": "伊丹",
    "伊丹": "伊丹",
    "大阪(関西)": "関空",
    "関西": "関空",
    "関空": "関空",
    "大阪(神戸)": "神戸",
    "神戸": "神戸",

    "札幌": "千歳",
    "札幌(新千歳)": "千歳",
    "新千歳": "千歳",
    "千歳": "千歳",
    "札幌(丘珠)": "丘珠",
    "丘珠": "丘珠",

    "名古屋": "中部",
    "名古屋(中部)": "中部",
    "中部": "中部",
    "名古屋(小牧)": "小牧",
    "小牧": "小牧",

    "沖縄(那覇)": "那覇",
    "沖縄": "那覇",
    "那覇": "那覇",

    "奄美大島": "奄美",
    "奄美": "奄美",
    "根室中標津": "根室",
    "中標津": "根室",
    "五島福江": "福江",
}


def normalize_svg_airport_name(name: str) -> str:
    """メール内の空港名を、SVG内 data-airport 名へ変換する。"""
    s = str(name or "").strip()
    s = s.replace(" ", "").replace("　", "")
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("*1", "").replace("*2", "")
    s = re.sub(r"空港$", "", s)
    return SVG_AIRPORT_ALIASES.get(s, s)


# ============================================================
# 時系列表示用データ
# ============================================================

@dataclass
class TimelineFlight:
    """テンプレート内 JavaScript に渡す1区間分のデータ。"""

    index: int
    boarding_date: str
    departure_time: str
    arrival_time: str
    duration: str
    departure: str
    arrival: str
    departure_svg: str
    arrival_svg: str
    flight_number: str
    fare_name: str
    seat_class: str
    seat_number: str
    ticket_number: str


def parse_date_for_sort(value: str) -> datetime:
    """搭乗日ソート用。解析できない値は最古日扱いにする。"""
    for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    return datetime(1900, 1, 1)


def parse_time_for_sort(value: str) -> datetime:
    """出発時刻ソート用。解析できない値は00:00扱いにする。"""
    try:
        return datetime.strptime(value or "00:00", "%H:%M")
    except ValueError:
        return datetime.strptime("00:00", "%H:%M")


def get_attr(row: object, name: str, default: str = "") -> str:
    """JALMailEngine の行オブジェクトから安全に文字列属性を取り出す。"""
    return str(getattr(row, name, default) or "")


def rows_to_timeline(rows: Iterable[object]) -> List[TimelineFlight]:
    """JALMailEngineの解析結果を、タイムライン表示用データへ変換する。"""
    flights: List[TimelineFlight] = []

    for row in rows:
        departure = get_attr(row, "departure")
        arrival = get_attr(row, "arrival")

        flights.append(
            TimelineFlight(
                index=0,
                boarding_date=get_attr(row, "boarding_date"),
                departure_time=get_attr(row, "departure_time"),
                arrival_time=get_attr(row, "arrival_time"),
                duration=get_attr(row, "duration"),
                departure=departure,
                arrival=arrival,
                departure_svg=normalize_svg_airport_name(departure),
                arrival_svg=normalize_svg_airport_name(arrival),
                flight_number=get_attr(row, "flight_number"),
                fare_name=get_attr(row, "fare_name"),
                seat_class=get_attr(row, "seat_class"),
                seat_number=get_attr(row, "seat_number"),
                ticket_number=get_attr(row, "ticket_number"),
            )
        )

    flights.sort(
        key=lambda f: (
            parse_date_for_sort(f.boarding_date),
            parse_time_for_sort(f.departure_time),
            f.flight_number,
            f.departure_svg,
            f.arrival_svg,
        )
    )

    for i, flight in enumerate(flights, start=1):
        flight.index = i

    return flights


# ============================================================
# JALMailEngine 呼び出し
# ============================================================

def create_jalmail_engine() -> object:
    """JALMailEngineのインスタンスを生成する。"""
    return JALMailEngine.open(
        IMAP_HOST,
        IMAP_PORT,
        IMAP_USER,
        IMAP_PASSWORD,
        IMAP_FOLDER,
        MAIL_SUBJECT_REGEX,
    )


def extract_timeline_from_imap(engine: object) -> List[TimelineFlight]:
    """IMAPから予約メールを取得してタイムラインデータを作る。"""
    rows = engine.get_rows_from_imap()
    return rows_to_timeline(rows)


def extract_timeline_from_text_files(engine: object, text_paths: Iterable[Path]) -> List[TimelineFlight]:
    """保存済みメール本文ファイルからタイムラインデータを作る。"""
    rows = engine.get_rows_from_files([str(path) for path in text_paths])
    return rows_to_timeline(rows)


# ============================================================
# テンプレート処理
# ============================================================

def load_base_network_html(path: Path) -> str:
    """HTMLテンプレートを読み込む。"""
    if not path.exists():
        raise FileNotFoundError(f"テンプレートファイルが見つかりません: {path}")
    return path.read_text(encoding="utf-8")



def build_logo_image_element(logo_path: str) -> str:
    """Build an SVG image element for an optional JPG logo overlay."""
    if not logo_path:
        return ""

    filename = Path(logo_path).name
    return f'''<image
      href="{filename}"
      x="70"
      y="90"
      width="430"
      height="410"
      preserveAspectRatio="xMidYMid meet"
      style="pointer-events:none"
    />'''

def inject_timeline_data(base_html: str, flights: List[TimelineFlight],
    logo_path: str = "",
) -> str:
    """テンプレートのプレースホルダへタイムラインJSONを差し込む。"""
    placeholder = "__TIMELINE_FLIGHTS_JSON__"
    if placeholder not in base_html:
        raise RuntimeError(f"テンプレート内に {placeholder} が見つかりません")

    flight_json = json.dumps([asdict(f) for f in flights], ensure_ascii=False, indent=2)

    logo_element = build_logo_image_element(logo_path)
    if LOGO_PLACEHOLDER in base_html:
        base_html = base_html.replace(LOGO_PLACEHOLDER, logo_element)

    return base_html.replace(placeholder, flight_json)


# 互換名: 以前の呼び出し名を残しておく
inject_timeline_ui = inject_timeline_data


# ============================================================
# 補助出力
# ============================================================

def write_debug_json(flights: List[TimelineFlight], output_path: Path) -> None:
    """抽出結果を確認用JSONとして出力する。"""
    output_path.write_text(
        json.dumps([asdict(f) for f in flights], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def print_summary(flights: List[TimelineFlight]) -> None:
    """抽出したフライト一覧を標準出力へ表示する。"""
    print(f"抽出区間数: {len(flights)}")
    for f in flights:
        print(
            f"#{f.index:02d} {f.boarding_date} {f.departure_time}-{f.arrival_time} "
            f"{f.flight_number} {f.departure_svg} -> {f.arrival_svg}"
        )


# ============================================================
# メイン
# ============================================================

def main() -> None:
    parser = argparse.ArgumentParser(description="JAL予約メールから予約経路チェック図HTMLを生成します")
    parser.add_argument("--output", default=OUTPUT_HTML, help="出力HTMLファイル")
    parser.add_argument("--template", default=BASE_NETWORK_TEMPLATE_PATH, help="jal_route_map.tmpl のパス")
    parser.add_argument("--logo", default="", help="SVG最前面へ表示する JPG ロゴ画像")
    parser.add_argument("--text", nargs="*", help="IMAPではなく、保存済みメール本文テキストを読むテスト用")
    parser.add_argument("--debug-json", default="", help="抽出した時系列データをJSON出力する場合のパス")
    parser.add_argument("--print-flights", action="store_true", help="抽出した時系列フライトを表示")
    args = parser.parse_args()

    print(f"IMAPサーバー: {IMAP_HOST}")
    print(f"IMAPフォルダ: {IMAP_FOLDER}")
    print("読取り中", end="")
    engine = create_jalmail_engine()

    if args.text:
        flights = extract_timeline_from_text_files(engine, [Path(path) for path in args.text])
    else:
        flights = extract_timeline_from_imap(engine)

    print("")
    if not flights:
        raise RuntimeError("フライト情報を抽出できませんでした")

    if args.print_flights:
        print_summary(flights)

    if args.debug_json:
        write_debug_json(flights, Path(args.debug_json))
        print(f"デバッグJSONを出力しました: {args.debug_json}")

    base_html = load_base_network_html(Path(args.template))
    output_html = inject_timeline_data(base_html, flights, logo_path=args.logo)
    Path(args.output).write_text(output_html, encoding="utf-8")
    print(f"HTMLを出力しました: {args.output}")


if __name__ == "__main__":
    main()
