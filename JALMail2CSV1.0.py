#
#   JAL航空券購入通知メールCSV変換ツール v1.0
#
#   Copyright(C) 2026 H.Kashima <kashima@kaele.com> Generated with Chat-GPT
#   2026.4.14   v1.0
#
import base64
import csv
import email
import imaplib
import math
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from email import policy
from email.header import decode_header
from typing import List, Optional, Tuple


# =========================
# IMAP設定
# =========================
IMAP_HOST = "imap.example.com"
IMAP_PORT = 993
IMAP_USER = "your_mail@example.com"
IMAP_PASSWORD = "your_password"
IMAP_FOLDER = "INBOX"

# 抽出する件名の正規表現
MAIL_SUBJECT_REGEX = r"^〔JAL国内線〕購入内容のお知らせ"

# =========================
# JALカードボーナス設定
# =========================
CARD_BONUS_TYPE = "なし"  # "なし", "普通", "CLUB-A以上"

# =========================
# ステータスボーナス設定
# 古い順に並べる
# =========================
STATUS_HISTORY = [
    ("一般", "2000/1/1"),
    # ("クリスタル", "2026/6/1"),
    # ("サファイア", "2026/9/1"),
    # ("ダイヤ", "2026/12/1"),
]
# =========================
# ツアープレミアム設定
# =========================
TOUR_PREMIUM = 0  # 1で有効

#
# ステータスボーナス積算率
#
STATUS_BONUS = {
    "一般": 0.0,
    "クリスタル": 0.55,
    "サファイア": 1.05,
    "ダイヤ": 1.3,
}

#
# ツアープレミアム適用運賃
#
TOUR_PREMIUM_TARGET = {
    "スペシャルセイバー",
    "個人包括旅行運賃",
}

#
# JALカード搭乗ボーナス積算率
#
CARD_BONUS_RULES = {
    "なし": {
        "first_flight_bonus": 0,
        "per_flight_rate": 0.0,
        "annual_cap": 0,
    },
    "普通": {
        "first_flight_bonus": 1000,
        "per_flight_rate": 0.10,
        "annual_cap": 5000,
    },
    "CLUB-A以上": {
        "first_flight_bonus": 2000,
        "per_flight_rate": 0.25,
        "annual_cap": 5000,
    },
}


# =========================
# FOP搭乗ボーナス表
# 運賃名は「乗継」を除いて評価
# =========================
FOP_BONUS_TABLE = [
    (
        400,
        {
            "フレックス",
            "JALカード割引",
            "ビジネスフレックス",
            "離島割引",
            "特定路線離島割引",
            "株主割引",
        },
    ),
    (
        200,
        {
            "セイバー",
            "スペシャルセイバー",
            "往復セイバー",
        },
    ),
    (
        0,
        {
            "プロモーション",
            "スカイメイト",
            "JALカードスカイメイト",
            "当日シニア割引",
            "特典航空券",
            "個人包括旅行運賃",
            "包括旅行運賃",
            "団体割引運賃",
            "学校研修割引運賃",
            "受注型企画旅行割引運賃",
            "募集型企画旅行割引運賃",
        },
    ),
]


# =========================
# 積算率表
# 運賃名は「乗継」を除いて評価
# =========================
MILE_RATE_TABLE = {
    "ファーストクラス": [
        (1.50, {"フレックス", "JALカード割引", "ビジネスフレックス"}),
        (1.25, {"セイバー", "往復セイバー", "株主割引"}),
        (1.00, {"スカイメイト", "JALカードスカイメイト", "当日シニア割引"}),
    ],
    "クラスＪ": [
        (1.10, {"フレックス", "JALカード割引", "ビジネスフレックス", "離島割引"}),
        (0.85, {"セイバー", "スペシャルセイバー", "往復セイバー", "株主割引"}),
        (0.60, {"スカイメイト", "JALカードスカイメイト", "当日シニア割引", "個人包括旅行運賃"}),
    ],
    "普通席": [
        (1.00, {"フレックス", "JALカード割引", "ビジネスフレックス", "離島割引", "特定路線離島割引"}),
        (0.75, {"セイバー", "スペシャルセイバー", "往復セイバー", "株主割引"}),
        (0.50, {"プロモーション", "当日シニア割引", "スカイメイト", "JALカードスカイメイト",
                "個人包括旅行運賃", "受注型企画旅行割引運賃", "学校研修割引運賃",
                "募集型企画旅行割引運賃"}),
    ],
}


# =========================
# 区間マイル表
# Web表に近い構造で管理
# =========================
ROUTE_MILES = {
    "東京": {
        "大阪": 280,
        "札幌": 510,
        "名古屋": 193,
        "福岡": 567,
        "沖縄(那覇)": 984,
        "女満別": 609,
        "旭川": 576,
        "釧路": 555,
        "帯広": 526,
        "函館": 424,
        "青森": 358,
        "三沢": 355,
        "秋田": 279,
        "花巻": 284,
        "仙台": 177,
        "山形": 190,
        "小松": 211,
        "南紀白浜": 303,
        "岡山": 356,
        "出雲": 405,
        "広島": 414,
        "山口宇部": 510,
        "徳島": 329,
        "高松": 354,
        "高知": 393,
        "松山": 438,
        "北九州": 534,
        "大分": 499,
        "長崎": 610,
        "熊本": 568,
        "宮崎": 561,
        "鹿児島": 601,
        "奄美大島": 787,
        "久米島": 1018,
        "宮古": 1158,
        "石垣": 1224,
    },
    "大阪": {
        "札幌": 666,
        "福岡": 287,
        "沖縄(那覇)": 739,
        "女満別": 797,
        "旭川": 739,
        "釧路": 753,
        "帯広": 711,
        "函館": 578,
        "青森": 523,
        "三沢": 536,
        "秋田": 439,
        "花巻": 474,
        "仙台": 396,
        "山形": 385,
        "新潟": 314,
        "松本": 183,
        "但馬": 68,
        "隠岐": 165,
        "出雲": 148,
        "松山": 159,
        "大分": 219,
        "長崎": 330,
        "熊本": 290,
        "宮崎": 292,
        "鹿児島": 329,
        "種子島": 379,
        "屋久島": 402,
        "奄美大島": 541,
        "徳之島": 603,
        "宮古": 906,
        "石垣": 969,
    },
    "神戸": {
        "青森": 523,
        "花巻": 474,
        "新潟": 314,
        "松本": 183,
        "高知": 119,
    },
    "札幌": {
        "名古屋": 614,
        "福岡": 882,
        "根室中標津": 178,
        "利尻": 159,
        "女満別": 148,
        "釧路": 136,
        "函館": 90,
        "奥尻": 123,
        "青森": 153,
        "三沢": 156,
        "秋田": 238,
        "花巻": 226,
        "仙台": 335,
        "山形": 321,
        "新潟": 369,
        "松本": 507,
        "静岡": 592,
        "出雲": 696,
        "広島": 749,
        "徳島": 715,
    },
    "函館": {
        "旭川": 162,
        "釧路": 194,
        "奥尻": 74,
        "三沢": 80,
    },
    "仙台": {
        "沖縄(那覇)": 1130,
        "出雲": 483,
    },
    "名古屋": {
        "沖縄(那覇)": 809,
        "釧路": 690,
        "帯広": 652,
        "青森": 465,
        "秋田": 380,
        "花巻": 409,
        "仙台": 322,
        "山形": 315,
        "新潟": 249,
        "出雲": 226,
        "高知": 201,
        "松山": 246,
        "福岡": 374,
        "北九州": 342,
        "長崎": 417,
        "熊本": 375,
        "鹿児島": 411,
        "宮古": 979,
        "石垣": 1044,
    },
    "出雲": {
        "静岡": 304,
        "神戸": 148,
        "隠岐": 65,
    },
    "静岡": {
        "北九州": 419,
        "熊本": 448,
    },
    "福岡": {
        "沖縄(那覇)": 537,
        "花巻": 724,
        "仙台": 665,
        "新潟": 572,
        "松本": 461,
        "静岡": 451,
        "出雲": 188,
        "徳島": 242,
        "高知": 187,
        "松山": 131,
        "天草": 78,
        "対馬": 81,
        "五島福江": 113,
        "宮崎": 131,
        "鹿児島": 125,
        "屋久島": 225,
        "奄美大島": 360,
    },
    "熊本": {
        "天草": 42,
    },
    "長崎": {
        "壱岐": 60,
        "五島福江": 67,
        "対馬": 98,
    },
    "宮崎": {
        "広島": 196,
    },
    "鹿児島": {
        "静岡": 479,
        "岡山": 268,
        "広島": 223,
        "高松": 260,
        "松山": 181,
        "種子島": 88,
        "屋久島": 102,
        "喜界島": 246,
        "奄美大島": 242,
        "徳之島": 296,
        "沖永良部": 326,
        "与論": 358,
    },
    "奄美大島": {
        "喜界島": 16,
        "徳之島": 65,
        "沖永良部": 92,
        "与論": 125,
    },
    "与論": {
        "沖永良部": 34,
    },
    "沖縄(那覇)": {
        "小松": 873,
        "岡山": 690,
        "松山": 607,
        "北九州": 563,
        "奄美大島": 199,
        "沖永良部": 107,
        "与論": 74,
        "北大東": 229,
        "南大東": 225,
        "久米島": 59,
        "宮古": 177,
        "石垣": 247,
        "与那国": 316,
    },
    "沖永良部": {
        "徳之島": 30,
    },
    "宮古": {
        "石垣": 72,
        "多良間": 39,
    },
    "石垣": {
        "与那国": 80,
    },
    "南大東": {
        "北大東": 8,
    },
}


# =========================
# データ構造
# =========================
@dataclass
class UnifiedRow:
    ticket_number: str
    issue_date: str
    fare_amount_yen: str
    boarding_date: str
    departure_time: str
    arrival_time: str
    duration: str
    departure: str
    arrival: str
    flight_number: str
    seat_class: str
    seat_number: str
    fare_name: str
    route_miles: Optional[int]
    flight_miles: Optional[int]
    card_bonus_miles: int
    bonus_fop: int
    accrued_miles: Optional[int]
    fop: Optional[int]
    remarks: str


# =========================
# ユーティリティ
# =========================
# 四捨五入
def round_half_up(value: float) -> int:
    return math.floor(value + 0.5)


# 日付の変換
def normalize_date(jp_date: str) -> str:
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", jp_date)
    if not m:
        return jp_date.strip()
    y, mo, d = m.groups()
    return f"{y}/{int(mo):02d}/{int(d):02d}"


# 空港名の変換
def normalize_airport(name: str) -> str:
    s = name.replace(" ", "").replace("　", "").strip()
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("*1", "").replace("*2", "")

    aliases = {
        "東京(羽田)": "東京",
        "東京(成田)": "東京",
        "大阪(伊丹)": "大阪",
        "大阪(関西)": "大阪",
        "大阪(神戸)": "神戸",
        "神戸": "神戸",
        "札幌(新千歳)": "札幌",
        "札幌(丘珠)": "札幌",
        "名古屋(中部)": "名古屋",
        "沖縄(那覇)": "沖縄(那覇)",
    }

    return aliases.get(s, s)


# 乗継の付く空港名の処理
def normalize_fare_name(fare_name: str) -> str:
    s = fare_name.strip()
    if s.endswith("乗継"):
        s = s[:-2].strip()
    return s


# 所要時間の計算
def calc_duration(dep: str, arr: str) -> str:
    if not dep or not arr:
        return ""
    try:
        t1 = datetime.strptime(dep, "%H:%M")
        t2 = datetime.strptime(arr, "%H:%M")
        if t2 < t1:
            # 通常は起こらない想定だが、日跨ぎ対策
            t2 = t2.replace(day=t2.day + 1)
        diff = t2 - t1
        minutes = int(diff.total_seconds() // 60)
        h = minutes // 60
        m = minutes % 60
        return f"{h:02d}:{m:02d}"
    except Exception:
        return ""


# 区間マイルの取得
def get_route_miles(dep: str, arr: str) -> int:
    dep_norm = normalize_airport(dep)
    arr_norm = normalize_airport(arr)

    if dep_norm in ROUTE_MILES and arr_norm in ROUTE_MILES[dep_norm]:
        return ROUTE_MILES[dep_norm][arr_norm]

    if arr_norm in ROUTE_MILES and dep_norm in ROUTE_MILES[arr_norm]:
        return ROUTE_MILES[arr_norm][dep_norm]

    print("エラー: ROUTE_MILESにエントリがありません")
    print(f"路線: {dep} → {arr}")
    sys.exit(1)


# 取得ステータスの確認
def get_status_for_boarding_date(boarding_date: str) -> Tuple[str, float]:
    try:
        target_date = datetime.strptime(boarding_date, "%Y/%m/%d")
    except Exception:
        return "一般", 0.0

    current_status = "一般"
    current_bonus = 0.0

    for status_name, acquired_date_str in STATUS_HISTORY:
        try:
            acquired_date = datetime.strptime(acquired_date_str, "%Y/%m/%d")
        except Exception:
            continue

        if target_date >= acquired_date:
            current_status = status_name
            current_bonus = STATUS_BONUS.get(status_name, 0.0)
        else:
            break

    return current_status, current_bonus


# 運賃から積算レートを取得
def get_mile_rate(seat_class: str, fare_name: str) -> float:
    seat_class = seat_class.strip()
    fare_name = normalize_fare_name(fare_name)

    rules = MILE_RATE_TABLE.get(seat_class, [])
    for rate, fare_set in rules:
        if fare_name in fare_set:
            return rate

    return 0.50


# FOPボーナスの取得
def get_fop_bonus(fare_name: str) -> int:
    fare_name = normalize_fare_name(fare_name)

    for bonus, fare_set in FOP_BONUS_TABLE:
        if fare_name in fare_set:
            return bonus

    return 0


# MIMEヘッダのデコード
def decode_mime_header(value: str) -> str:
    if not value:
        return ""

    parts = decode_header(value)
    decoded_parts = []

    for part, enc in parts:
        if isinstance(part, bytes):
            for charset in [enc, "utf-8", "cp932", "iso-2022-jp", "latin-1"]:
                if not charset:
                    continue
                try:
                    decoded_parts.append(part.decode(charset, errors="replace"))
                    break
                except Exception:
                    continue
            else:
                decoded_parts.append(part.decode("utf-8", errors="replace"))
        else:
            decoded_parts.append(part)

    return "".join(decoded_parts)


# textのデコード
def decode_text_payload(part) -> str:
    payload = part.get_payload(decode=True)
    if payload is None:
        raw = part.get_payload()
        return raw if isinstance(raw, str) else ""

    charset = part.get_content_charset()
    for cs in [charset, "utf-8", "cp932", "iso-2022-jp", "euc-jp", "latin-1"]:
        if not cs:
            continue
        try:
            return payload.decode(cs, errors="replace")
        except Exception:
            pass

    return payload.decode("utf-8", errors="replace")


# plain/textパートの読取り
def extract_text_plain(msg) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = (part.get_content_disposition() or "").lower()

            if content_type == "text/plain" and disposition != "attachment":
                text = decode_text_payload(part)
                if text.strip():
                    return text
        return ""

    if msg.get_content_type() == "text/plain":
        return decode_text_payload(msg)

    return ""


# utf-8 IMAPフォルダ名のエンコード
def encode_imap_folder_name(name: str) -> str:
    """
    IMAPフォルダ名をModified UTF-7に変換
    """
    res = []
    buf = ""

    for ch in name:
        if 0x20 <= ord(ch) <= 0x7E and ch != "&":
            if buf:
                b64 = base64.b64encode(buf.encode("utf-16be")).decode("ascii").rstrip("=")
                b64 = b64.replace("/", ",")
                res.append("&" + b64 + "-")
                buf = ""
            res.append(ch)
        elif ch == "&":
            if buf:
                b64 = base64.b64encode(buf.encode("utf-16be")).decode("ascii").rstrip("=")
                b64 = b64.replace("/", ",")
                res.append("&" + b64 + "-")
                buf = ""
            res.append("&-")
        else:
            buf += ch

    if buf:
        b64 = base64.b64encode(buf.encode("utf-16be")).decode("ascii").rstrip("=")
        b64 = b64.replace("/", ",")
        res.append("&" + b64 + "-")

    return "".join(res)


# =========================
# ノイズ除去
# =========================
def is_noise_line_mail(line: str) -> bool:
    noise_patterns = [
        r"^いつもJALグループをご利用いただきありがとうございます。$",
        r"^以下のご予約のお支払いを承りましたので確認してください。$",
        r"^〔運賃額変更時の取り扱い〕$",
        r"^〔ご搭乗方法〕$",
        r"^〔予約確認〕$",
        r"^■領収書$",
        r"^JAL Webサイト（eチケット検索）より電子領収書を発行いただけます。$",
        r"^〔eチケット検索〕$",
        r"^■おすすめ情報$",
        r"^〔ホテル・レンタカーなど〕$",
        r"^ホテルやレンタカーなどの予約でも、マイルがたまっておトクです。$",
        r"^〔旅のサポート〕$",
        r"^保険、タクシー、地上交通、オプショナルツアー、駐車場などのご利用でもマイルがたまります。$",
        r"^〔手荷物宅配・配送サービス〕$",
        r"^お客さまのニーズに合わせた手荷物宅配・配送サービスを実施しております。$",
        r"^■JAL国内線お問い合わせ窓口$",
        r"^〔お客さまサポート〕$",
        r"^当メールは送信専用です。.*$",
        r"^〔お心当たりがないメールが届いた場合のお手続きについて〕$",
        r"^日本航空\s+https://www\.jal\.co\.jp/.*$",
        r"^https://.*$",
        r"^※運航ダイヤ.*$",
        r"^※運賃情報には.*$",
        r"^■ご搭乗案内$",
        r"^ご搭乗日当日は、以下のいずれかをお持ちください。$",
        r"^・JALカード.*$",
        r"^・搭乗用2次元バーコード.*$",
        r"^・おサイフケータイ.*$",
        r"^・航空券番号$",
        r"^※「おサイフケータイ」.*$",
    ]
    s = line.strip()
    if not s:
        return True
    return any(re.search(p, s) for p in noise_patterns)


# =========================
# メール取得
# =========================
def fetch_matching_mail_texts() -> List[str]:
    matched_texts: List[str] = []

    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as mail:
        mail.login(IMAP_USER, IMAP_PASSWORD)

        encoded_folder = encode_imap_folder_name(IMAP_FOLDER)
        status, _ = mail.select(f'"{encoded_folder}"')
        if status != "OK":
            raise RuntimeError(f"IMAPフォルダを開けません: {IMAP_FOLDER}")

        status, data = mail.search(None, "ALL")
        if status != "OK":
            raise RuntimeError("IMAP検索に失敗しました")

        msg_ids = data[0].split()
        subject_re = re.compile(MAIL_SUBJECT_REGEX)

        for msg_id in msg_ids:
            print(".", end="", flush=True)

            status, msg_data = mail.fetch(msg_id, "(BODY.PEEK[])")
            if status != "OK":
                continue

            raw_email = None
            for item in msg_data:
                if isinstance(item, tuple) and len(item) >= 2 and item[1]:
                    raw_email = item[1]
                    break

            if raw_email is None:
                continue

            msg = email.message_from_bytes(raw_email, policy=policy.default)
            subject = decode_mime_header(msg.get("Subject", ""))

            if not subject_re.search(subject):
                continue

            text = extract_text_plain(msg)
            if text.strip():
                matched_texts.append(text)

    print()
    return matched_texts


# =========================
# マイル・FOP計算
# =========================
def calc_fop(
    route_miles: Optional[int],
    seat_class: str,
    fare_name: str,
    boarding_date: str,
) -> Tuple[Optional[int], Optional[int], int, Optional[int], str]:
    normal_fare_name = normalize_fare_name(fare_name)

    # 通常積算率（FOP用）
    base_rate = get_mile_rate(seat_class, normal_fare_name)

    # FOPボーナス
    bonus = get_fop_bonus(normal_fare_name)

    if route_miles is None:
        return None, None, bonus, None, ""

    remarks = []

    # ステータスボーナス
    status_name, bonus_mile = get_status_for_boarding_date(boarding_date)

    # =========================
    # フライトマイル（FOP基準）
    # =========================
    flight_miles = round_half_up(route_miles * base_rate)

    # =========================
    # 積算マイル用レート
    # TOUR_PREMIUMは積算マイルにのみ適用
    # =========================
    accrual_rate = base_rate

    if TOUR_PREMIUM == 1 and normal_fare_name in TOUR_PREMIUM_TARGET:
        accrual_rate = 1.0
        tp_bonus_rate = int(round((accrual_rate - base_rate) * 100))
        if tp_bonus_rate > 0:
            remarks.append(f"TP{tp_bonus_rate}")

    accrual_base_miles = round_half_up(route_miles * accrual_rate)

    # =========================
    # 積算マイル
    # =========================
    total_miles = accrual_base_miles + (accrual_base_miles * bonus_mile)

    if bonus_mile > 0:
        status_code_map = {
            "クリスタル": "ク",
            "サファイア": "サ",
            "ダイヤ": "ダ",
        }
        status_code = status_code_map.get(status_name, "")
        if status_code:
            remarks.append(f"{status_code}{int(round(bonus_mile * 100))}")

    accrued = math.ceil(total_miles)

    # =========================
    # FOP（TOUR_PREMIUM非適用）
    # フライトマイル × 2 + ボーナスFOP
    # =========================
    fop = flight_miles * 2 + bonus

    return flight_miles, accrued, bonus, fop, ",".join(remarks)


# =========================
# メール本文解析
# =========================
def parse_mail_text(text: str) -> List[UnifiedRow]:
    raw_lines = [line.strip() for line in text.splitlines()]
    lines = [line for line in raw_lines if not is_noise_line_mail(line)]

    rows: List[UnifiedRow] = []

    current_ticket_number = ""
    current_issue_date = ""
    current_fare_amount = ""
    current_fare_name = ""
    current_fare_name_raw = ""

    # -------------------------
    # 1st pass: 共通情報を先に集める
    # -------------------------
    for line in lines:
        m_issue = re.match(r"^■ご予約内容（(\d{4}年\d{1,2}月\d{1,2}日).*$", line)
        if m_issue:
            current_issue_date = normalize_date(m_issue.group(1))
            continue

        m_ticket = re.match(r"^航空券番号：\s*(\d+)$", line)
        if m_ticket:
            current_ticket_number = m_ticket.group(1)
            continue

        m_fare = re.match(r"^(.+?)\s+大人：\d+名×[\d,]+円$", line)
        if m_fare:
            raw_fare_name = m_fare.group(1).strip()
            candidate_fare_name = normalize_fare_name(raw_fare_name)

            found = False
            for seat_rules in MILE_RATE_TABLE.values():
                for _, fare_set in seat_rules:
                    if candidate_fare_name in fare_set:
                        found = True
                        break
                if found:
                    break

            if not found and candidate_fare_name in TOUR_PREMIUM_TARGET:
                found = True

            if found:
                current_fare_name = candidate_fare_name
                current_fare_name_raw = raw_fare_name
            continue

        m_total = re.match(r"^合計\s+([\d,]+)円$", line)
        if m_total:
            current_fare_amount = m_total.group(1).replace(",", "")
            continue

    # -------------------------
    # 2nd pass: フライト明細を作る
    # -------------------------
    i = 0
    while i < len(lines):
        line = lines[i]

        m_flight_header = re.match(r"^(\d{4}年\d{1,2}月\d{1,2}日).*?(JAL|JTA|RAC)(\d+)便$", line)
        if m_flight_header:
            boarding_date = normalize_date(m_flight_header.group(1))
            flight_number = f"{m_flight_header.group(2)}{m_flight_header.group(3)}"

            departure_time = ""
            arrival_time = ""
            departure = ""
            arrival = ""
            seat_class = ""
            seat_number = ""

            if i + 1 < len(lines):
                m_route = re.match(
                    r"^(.*?)\s*(\d{1,2}:\d{2})発\s*→\s*(.*?)\s*(\d{1,2}:\d{2})着$",
                    lines[i + 1]
                )
                if m_route:
                    departure = m_route.group(1).strip()
                    departure_time = m_route.group(2)
                    arrival = m_route.group(3).strip()
                    arrival_time = m_route.group(4)

            j = i + 2
            while j < len(lines):
                if re.match(r"^\(旅程\d+\)$", lines[j]):
                    break
                if re.match(r"^\d{4}年\d{1,2}月\d{1,2}日.*?(JAL|JTA|RAC)\d+便$", lines[j]):
                    break
                if lines[j].startswith("〔運賃情報〕"):
                    break

                m_seat_class = re.match(r"^座席：(.+)$", lines[j])
                if m_seat_class:
                    seat_class = m_seat_class.group(1).strip()

                m_seat_number = re.match(r"^座席番号：(.+)$", lines[j])
                if m_seat_number:
                    seat_number = m_seat_number.group(1).strip()

                j += 1

            duration = calc_duration(departure_time, arrival_time)
            route_miles = get_route_miles(departure, arrival)
            flight_miles, accrued_miles, bonus_fop, fop, remarks = calc_fop(
                route_miles,
                seat_class,
                current_fare_name,
                boarding_date,
            )

            fare_name = current_fare_name_raw

            rows.append(
                UnifiedRow(
                    ticket_number=current_ticket_number,
                    issue_date=current_issue_date,
                    fare_amount_yen=current_fare_amount,
                    boarding_date=boarding_date,
                    departure_time=departure_time,
                    arrival_time=arrival_time,
                    duration=duration,
                    departure=departure,
                    arrival=arrival,
                    flight_number=flight_number,
                    seat_class=seat_class,
                    seat_number=seat_number,
                    fare_name=fare_name,
                    route_miles=route_miles,
                    flight_miles=flight_miles,
                    card_bonus_miles=0,
                    bonus_fop=bonus_fop,
                    accrued_miles=accrued_miles,
                    fop=fop,
                    remarks=remarks,
                )
            )

            i = j
            continue

        i += 1

    return rows


# =========================
# JALカードボーナス適用
# =========================
def apply_card_bonus(rows: List[UnifiedRow]) -> None:
    rule = CARD_BONUS_RULES.get(CARD_BONUS_TYPE, CARD_BONUS_RULES["なし"])

    first_flight_bonus = rule["first_flight_bonus"]
    per_flight_rate = rule["per_flight_rate"]
    annual_cap = rule["annual_cap"]

    if annual_cap <= 0:
        return

    yearly_bonus_used = {}
    yearly_first_applied = {}

    rows.sort(key=lambda r: (
        datetime.strptime(r.boarding_date, "%Y/%m/%d"),
        datetime.strptime(r.departure_time or "00:00", "%H:%M"),
    ))

    for row in rows:
        try:
            year = datetime.strptime(row.boarding_date, "%Y/%m/%d").year
        except Exception:
            continue

        if year not in yearly_bonus_used:
            yearly_bonus_used[year] = 0
            yearly_first_applied[year] = False

        remaining = annual_cap - yearly_bonus_used[year]
        if remaining <= 0:
            row.card_bonus_miles = 0
            continue

        bonus_to_add = 0
        remark_list = [x for x in row.remarks.split(",") if x]

        # 年初回搭乗ボーナス
        if not yearly_first_applied[year]:
            bonus_to_add += first_flight_bonus
            yearly_first_applied[year] = True
            if first_flight_bonus > 0:
                remark_list.append("年初")

        # 搭乗ごとのボーナス
        if row.route_miles is not None:
            rate = get_mile_rate(row.seat_class, row.fare_name)

            # カードボーナス基準は区間マイル × 運賃積算率
            base_miles = round_half_up(row.route_miles * rate)

            per_flight_bonus = round_half_up(base_miles * per_flight_rate)
            bonus_to_add += per_flight_bonus

            if per_flight_rate > 0:
                remark_list.append(f"搭{int(round(per_flight_rate * 100))}")

        # 年間上限
        if bonus_to_add > remaining:
            bonus_to_add = remaining

        row.card_bonus_miles = bonus_to_add

        if row.accrued_miles is None:
            row.accrued_miles = bonus_to_add
        else:
            row.accrued_miles += bonus_to_add

        yearly_bonus_used[year] += bonus_to_add
        row.remarks = ",".join(dict.fromkeys(remark_list))


# =========================
# CSV出力
# =========================
def write_csv(rows: List[UnifiedRow], output_path: str):
    fieldnames = [
        "航空券番号",
        "発行日",
        "運賃額(円)",
        "搭乗日",
        "出発時刻",
        "到着時刻",
        "所要時間",
        "出発",
        "到着",
        "便名",
        "座席",
        "座席番号",
        "運賃",
        "区間マイル",
        "フライトマイル",
        "JALカード",
        "ボーナスFOP",
        "積算マイル",
        "FOP",
        "備考",
    ]

    seen_ticket_numbers = set()

    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        for r in rows:
            if r.ticket_number and r.ticket_number not in seen_ticket_numbers:
                fare_amount = r.fare_amount_yen
                seen_ticket_numbers.add(r.ticket_number)
            else:
                fare_amount = ""

            writer.writerow({
                "航空券番号": r.ticket_number,
                "発行日": r.issue_date,
                "運賃額(円)": fare_amount,
                "搭乗日": r.boarding_date,
                "出発時刻": r.departure_time,
                "到着時刻": r.arrival_time,
                "所要時間": r.duration,
                "出発": r.departure,
                "到着": r.arrival,
                "便名": r.flight_number,
                "座席": r.seat_class,
                "座席番号": r.seat_number,
                "運賃": r.fare_name,
                "区間マイル": r.route_miles if r.route_miles is not None else "",
                "フライトマイル": r.flight_miles if r.flight_miles is not None else "",
                "JALカード": r.card_bonus_miles,
                "ボーナスFOP": r.bonus_fop,
                "積算マイル": r.accrued_miles if r.accrued_miles is not None else "",
                "FOP": r.fop if r.fop is not None else "",
                "備考": r.remarks,
            })


# =========================
# 集計表示
# =========================
def print_summary(rows: List[UnifiedRow], mail_count: int):
    yearly = {}

    for r in rows:
        try:
            year = datetime.strptime(r.boarding_date, "%Y/%m/%d").year
        except Exception:
            year = "不明"

        if year not in yearly:
            yearly[year] = {
                "segments": 0,
                "flight_miles": 0,
                "card_bonus": 0,
                "miles": 0,
                "fop": 0,
            }

        yearly[year]["segments"] += 1
        yearly[year]["flight_miles"] += r.flight_miles or 0
        yearly[year]["card_bonus"] += r.card_bonus_miles or 0
        yearly[year]["miles"] += r.accrued_miles or 0
        yearly[year]["fop"] += r.fop or 0

    print(f"メール件数: {mail_count}")
    print(f"区間数合計: {len(rows)}")

    for year in sorted(yearly, key=lambda x: (str(x))):
        print(
            f"{year}年: "
            f"区間数={yearly[year]['segments']}, "
            f"フライトマイル合計={yearly[year]['flight_miles']}, "
            f"JALカードボーナス合計={yearly[year]['card_bonus']}, "
            f"積算マイル合計={yearly[year]['miles']}, "
            f"FOP合計={yearly[year]['fop']}"
        )


# =========================
# 実行
# =========================
if __name__ == "__main__":
    output_path = "jal_tickets_with_fop.csv"

    mail_texts = fetch_matching_mail_texts()
    if not mail_texts:
        raise RuntimeError("条件に一致するメールが見つかりませんでした")

    all_rows: List[UnifiedRow] = []
    for mail_text in mail_texts:
        rows = parse_mail_text(mail_text)
        all_rows.extend(rows)

    apply_card_bonus(all_rows)
    write_csv(all_rows, output_path)
    print_summary(all_rows, len(mail_texts))

    print(f"CSVを出力しました: {output_path}")
