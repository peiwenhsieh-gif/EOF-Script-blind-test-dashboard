#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]

DEFAULT_OUTPUTS = {
    "dashboard_js": ROOT / "blind-test-dashboard-data-0401.js",
    "open_response_js": ROOT / "blind-test-open-response-data.js",
    "workbook_rows_json": ROOT / "workbook_rows_0401.json",
    "analysis_js": ROOT / "workbook-open-response-rows.js",
}

DEFAULT_SECTIONS = [
    {"id": "audience", "label": "觀眾輪廓", "title": "觀眾輪廓", "subtitle": "背景輪廓、類型偏好與近期觀影"},
    {"id": "worldview", "label": "初始認知", "title": "世界觀初始認知", "subtitle": "妖怪傳說與核彈歷史知曉比例"},
    {"id": "characters", "label": "角色分析", "title": "角色分析", "subtitle": "角色喜好、續集主導角色與原話"},
    {"id": "powers", "label": "超能力", "title": "超能力評估", "subtitle": "最喜歡 / 最無聊 / 看不懂與能力矩陣"},
    {"id": "plot", "label": "劇情理解", "title": "場面 / 劇情理解", "subtitle": "折線圖、關係線與合理性"},
    {"id": "reception", "label": "市場反應", "title": "市場反應", "subtitle": "聯想象限、劇本評分、入場 / 推薦與原因"},
    {"id": "cross", "label": "交叉分析", "title": "交叉分析", "subtitle": "性別分析與年齡分析"},
    {"id": "qualitative", "label": "質化回饋", "title": "質化回饋", "subtitle": "雙欄主題分群與原始回覆"},
]

DEFAULT_PALETTE = {
    "role": {
        "虎姑婆": "var(--c-tiger)",
        "白賊七": "var(--c-bai)",
        "好鼻師": "var(--c-nose)",
        "小潔": "var(--c-jie)",
        "局長": "var(--c-boss)",
        "水蛇": "var(--c-boss-soft)",
        "蜥蜴": "var(--c-boss-lite)",
        "水男孩": "var(--c-boss-lite)",
    },
    "gender": {
        "男": "var(--male)",
        "女": "var(--female)",
    },
}

DEFAULT_ASSOCIATION_ZONES = [
    {"text": "歐美超英宇宙", "x": 8, "y": 8},
    {"text": "亞洲主流動漫", "x": 74, "y": 8},
    {"text": "西方奇幻科幻", "x": 8, "y": 84},
    {"text": "亞洲獨立動漫", "x": 72, "y": 84},
]

QUESTION_MAP = {
    "姓名": "name",
    "性別": "gender",
    "年齡": "age",
    "目前居住縣市": "location",
    "工作產業 / 學生請填就讀科系": "occupation",
    "進「電影院」觀影頻率": "frequency",
    "最近進電影院看的電影（數量不限）": "recentMovies",
    "喜歡的電影類型 （多選）": "genre",
    "閱讀劇本前，知道下列哪些妖怪傳說？（可複選）": "folklore",
    "閱讀劇本前，是否知道台灣有曾經製作過「核彈」的這段歷史？": "nukeAware",
    "你認為主角是誰？(單選)": "hero",
    "你喜歡的角色？ (可複選)": "likes",
    "承上題，你喜歡這些角色的原因？": "likeReason",
    "你最不喜歡的角色？(單選)": "dislike",
    "承上題，你不喜歡這個角色的原因？": "dislikeReason",
    "下列哪個妖怪的超能力對你來說不清楚、看不懂？(可複選)": "confusing",
    "下列哪個妖怪的超能力對你來說，覺得最無聊、最普通？(單選)": "boringPower",
    "下列哪個妖怪的超能力對你來說，覺得最喜歡、最酷？(單選)": "coolPower",
    "你覺得「局長」的陰謀是什麼？他為什麼要計劃這個陰謀？": "villainResponse",
    "閱讀劇本後，你最期待被拍攝出來的場面？或是在文字描繪上有讓你產生強烈的畫面感？\n請提供頁數/場次或具體情節。": "expectedResponse",
    "你是否有看不懂的情節/橋段？\n請說明頁數/場次或具體情節，以及你不理解的部分。": "confusingResponse",
    "劇本哪一段的情節讓你覺得無聊或拖沓？\n請說明頁數/場次或具體情節，以及你覺得拖沓的原因。": "boringResponse",
    "你是否覺得劇本有哪些地方不合理/出戲的地方？\n請說明頁數/場次或具體情節，以及你覺得不合理/出戲的地方。": "unreasonableResponse",
    "你在劇本哪一場次，知道虎哥跟虎姑婆是情侶關係？(單選)": "tigerReveal",
    "你在劇本哪一場次，知道水蛇跟水男孩是母子關係？(單選)": "motherReveal",
    "你覺得覺得小潔跟虎姑婆在哪一場次開始產生革命情感？(單選)": "emotionReveal",
    "在第113場，看到護理師出現在後山基地時，你是否會覺得不合理？\n請回答合理/不合理，並說明原因。": "nurseResponse",
    "閱讀劇本時，你是否有聯想到其他電影? 請舉例說明": "association",
    "劇本喜愛程度": "scriptScore",
    "請用一句話說明，你覺得這個劇本在講述一個什麼故事？": "logline",
    "閱讀劇本後，你是否會想進電影院觀看本片？": "entrance",
    "你是否願意推薦本片給親朋好友？": "recommend",
    "承上題，你會推薦本片的原因？(可複選)": "recommendReasons",
    "看完劇本後，你會期待這個「變形者宇宙 」的下一部續集或外傳嗎？會期待由哪個角色擔任主角？": "sequel",
    "你是否還有任何想要補充或分享的想法？": "finalNote",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Rebuild dashboard data files from updated survey and analysis Excel workbooks.")
    parser.add_argument("--survey", type=Path, help="Path to the updated survey response workbook (.xlsx).")
    parser.add_argument("--analysis", type=Path, help="Path to the updated qualitative analysis workbook (.xlsx).")
    parser.add_argument("--root", type=Path, default=ROOT, help="Project root. Defaults to the current repository.")
    return parser.parse_args()


def find_latest_xlsx(root: Path, patterns: list[str]) -> Path:
    matches: list[Path] = []
    for pattern in patterns:
        matches.extend(root.glob(pattern))
    matches = [path for path in matches if path.is_file()]
    if not matches:
        raise FileNotFoundError(f"Could not find workbook matching any of: {', '.join(patterns)}")
    return max(matches, key=lambda path: path.stat().st_mtime)


def col_to_index(col_ref: str) -> int:
    value = 0
    for char in col_ref:
        if char.isalpha():
            value = value * 26 + (ord(char.upper()) - 64)
    return max(value - 1, 0)


def parse_xlsx(path: Path) -> dict[str, list[list[str]]]:
    namespace = {
        "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    with zipfile.ZipFile(path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", namespace):
                shared_strings.append("".join(node.text or "" for node in item.iterfind(".//a:t", namespace)))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {node.attrib["Id"]: node.attrib["Target"] for node in rel_root}

        sheets: dict[str, list[list[str]]] = {}
        for sheet in workbook.find("a:sheets", namespace) or []:
            rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = "xl/" + rel_map[rel_id]
            sheet_root = ET.fromstring(archive.read(target))
            sheet_data = sheet_root.find("a:sheetData", namespace)
            rows: list[list[str]] = []
            if sheet_data is None:
                sheets[sheet.attrib["name"]] = rows
                continue
            for row in sheet_data:
                values: list[str] = []
                for cell in row:
                    ref = "".join(ch for ch in cell.attrib.get("r", "") if ch.isalpha())
                    index = col_to_index(ref) if ref else len(values)
                    while len(values) < index:
                        values.append("")
                    cell_type = cell.attrib.get("t")
                    value = ""
                    if cell_type == "inlineStr":
                        value = "".join(node.text or "" for node in cell.iterfind(".//a:t", namespace))
                    else:
                        raw_node = cell.find("a:v", namespace)
                        if raw_node is not None:
                            raw_value = raw_node.text or ""
                            value = shared_strings[int(raw_value)] if cell_type == "s" else raw_value
                    values.append(value)
                rows.append(values)
            sheets[sheet.attrib["name"]] = rows
        return sheets


def tidy_text(value: Any) -> str:
    return str(value or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def split_multi(value: Any) -> list[str]:
    text = tidy_text(value)
    if not text:
        return []
    parts = re.split(r"\s*(?:,|;|；|\n)+\s*", text)
    return [part.strip() for part in parts if part.strip()]


def normalize_gender(value: str) -> str:
    text = tidy_text(value)
    if "男" in text:
        return "男"
    if "女" in text:
        return "女"
    return text


def normalize_location(value: str) -> str:
    text = tidy_text(value).replace("臺", "台")
    return text or "其他"


def normalize_bool(value: str) -> bool:
    text = tidy_text(value)
    return text in {"是", "會", "願意", "true", "TRUE", "True", "1"}


def normalize_int(value: Any) -> int:
    text = tidy_text(value)
    if not text:
        return 0
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return 0
    return int(round(float(match.group(0))))


def normalize_movie_title(raw_title: str) -> str:
    title = tidy_text(raw_title)
    title = re.sub(r"\s+", " ", title)
    if "前幾天看什麼女孩的" in title:
        return "女孩"
    canonical_rules = [
        (r"動物方城市\s*2|動物方城市2", "動物方城市 2"),
        (r"鬼滅(?:之刃)?無限城(?:篇)?", "鬼滅無限城"),
        (r"鏈鋸人\s*蕾潔篇|鏈鋸人蕾潔篇", "鏈鋸人 蕾潔篇"),
        (r"不可能的任務[:：]?\s*最終清算|不可能的任務最終清算", "不可能的任務"),
        (r"陽光(?:女子|少女)合唱團|陽光女子$", "陽光女子合唱團"),
        (r"阿凡達\s*3|阿凡達3|阿凡達三|阿煩答\s*火之燼", "阿凡達3"),
        (r"魔法壞女巫[:：]?\s*第二部|魔法壞女巫2", "魔法壞女巫"),
    ]
    for pattern, canonical in canonical_rules:
        if re.search(pattern, title, re.I):
            return canonical
    return title


def expand_recent_movie_titles(raw_title: str) -> list[str]:
    title = tidy_text(raw_title)
    if not title:
        return []
    normalized = re.sub(r"\s+", " ", title)
    if normalized == "阿凡達 出神入化3":
        return ["阿凡達3", "出神入化3"]
    canonical = normalize_movie_title(normalized)
    return [canonical] if canonical else []


def classify_movie_group(title: str) -> str:
    if re.search(r"(陽光女子合唱團|大濛|國寶|雙囍|女孩|周處|鬼才之道|夜校女生|冠軍之路|那張照片裡的我們)", title):
        return "台灣國片"
    if re.search(r"(動物方城市|動物方程式|鬼滅|鏈鋸人|藤本樹|mygo|荒野機器人|排球少年|名偵探柯南|蠟筆小新|魔法公主)", title, re.I):
        return "動畫 / 日漫"
    if re.search(r"(阿凡達|不可能的任務|魔法壞女巫|F1|超人|侏羅紀|漫威|dc|玩命關頭|007|全知讀者視角|出神入化)", title, re.I):
        return "好萊塢 / 視效大片"
    return "劇情 / 其他"


def derive_recent_movie_groups(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    group_order = [
        ("台灣國片", "var(--gold)"),
        ("好萊塢 / 視效大片", "var(--orange)"),
        ("動畫 / 日漫", "var(--blue)"),
        ("劇情 / 其他", "var(--green)"),
    ]
    counters: dict[str, Counter[str]] = {group: Counter() for group, _ in group_order}
    for row in rows:
        raw_value = tidy_text(row.get("recentMovies"))
        if not raw_value:
            continue
        candidates = re.split(r"\s*(?:,|，|/|、|\n)+\s*", raw_value)
        for candidate in candidates:
            for title in expand_recent_movie_titles(candidate):
                if not title:
                    continue
                counters[classify_movie_group(title)][title] += 1

    groups = []
    for group_name, color in group_order:
        items = [{"label": label, "count": count} for label, count in counters[group_name].most_common()]
        groups.append({"title": group_name, "color": color, "items": items})
    return groups


def rows_from_sheet(sheet_rows: list[list[str]]) -> list[dict[str, str]]:
    if not sheet_rows:
        return []
    headers = [tidy_text(value) for value in sheet_rows[0]]
    normalized_rows: list[dict[str, str]] = []
    for row in sheet_rows[1:]:
        padded = row + [""] * max(0, len(headers) - len(row))
        item = {headers[index]: tidy_text(padded[index]) for index in range(len(headers)) if headers[index]}
        if any(value for value in item.values()):
            normalized_rows.append(item)
    return normalized_rows


def choose_survey_sheet(sheets: dict[str, list[list[str]]]) -> list[list[str]]:
    for preferred in ("表單回覆 整合", "表單回覆 整合-公式"):
        if preferred in sheets:
            return sheets[preferred]
    for rows in sheets.values():
        if rows and "姓名" in rows[0]:
            return rows
    raise ValueError("Could not find a usable survey sheet containing the expected headers.")


def map_survey_row(raw_row: dict[str, str]) -> dict[str, Any]:
    mapped: dict[str, Any] = {}
    for source_key, target_key in QUESTION_MAP.items():
        mapped[target_key] = raw_row.get(source_key, "")

    respondent = {
        "name": tidy_text(mapped["name"]),
        "gender": normalize_gender(mapped["gender"]),
        "age": tidy_text(mapped["age"]),
        "location": normalize_location(mapped["location"]),
        "occupation": tidy_text(mapped["occupation"]),
        "frequency": tidy_text(mapped["frequency"]),
        "recentMovies": tidy_text(mapped["recentMovies"]),
        "genre": tidy_text(mapped["genre"]),
        "folklore": split_multi(mapped["folklore"]),
        "nukeAware": normalize_bool(mapped["nukeAware"]),
        "hero": tidy_text(mapped["hero"]),
        "likes": split_multi(mapped["likes"]),
        "likeReason": tidy_text(mapped["likeReason"]),
        "dislike": tidy_text(mapped["dislike"]),
        "dislikeReason": tidy_text(mapped["dislikeReason"]),
        "confusing": split_multi(mapped["confusing"]),
        "boringPower": tidy_text(mapped["boringPower"]),
        "coolPower": tidy_text(mapped["coolPower"]),
        "villainResponse": tidy_text(mapped["villainResponse"]),
        "expectedResponse": tidy_text(mapped["expectedResponse"]),
        "confusingResponse": tidy_text(mapped["confusingResponse"]),
        "boringResponse": tidy_text(mapped["boringResponse"]),
        "unreasonableResponse": tidy_text(mapped["unreasonableResponse"]),
        "tigerReveal": tidy_text(mapped["tigerReveal"]),
        "motherReveal": tidy_text(mapped["motherReveal"]),
        "emotionReveal": tidy_text(mapped["emotionReveal"]),
        "nurseResponse": tidy_text(mapped["nurseResponse"]),
        "association": tidy_text(mapped["association"]),
        "scriptScore": normalize_int(mapped["scriptScore"]),
        "logline": tidy_text(mapped["logline"]),
        "entrance": normalize_bool(mapped["entrance"]),
        "recommend": normalize_bool(mapped["recommend"]),
        "recommendReasons": split_multi(mapped["recommendReasons"]),
        "sequel": tidy_text(mapped["sequel"]),
        "finalNote": tidy_text(mapped["finalNote"]),
    }
    return respondent


def build_reason_data(respondents: list[dict[str, Any]]) -> dict[str, dict[str, list[str]]]:
    like_buckets: dict[str, list[str]] = defaultdict(list)
    dislike_buckets: dict[str, list[str]] = defaultdict(list)
    for row in respondents:
        like_reason = tidy_text(row.get("likeReason"))
        if like_reason:
            for role in row.get("likes", []):
                if role:
                    like_buckets[role].append(like_reason)
        dislike_reason = tidy_text(row.get("dislikeReason"))
        dislike_role = tidy_text(row.get("dislike"))
        if dislike_role and dislike_reason:
            dislike_buckets[dislike_role].append(dislike_reason)
    return {
        "like": dict(sorted(like_buckets.items(), key=lambda item: item[0])),
        "dislike": dict(sorted(dislike_buckets.items(), key=lambda item: item[0])),
    }


def build_open_response_data(respondents: list[dict[str, Any]]) -> dict[str, list[str]]:
    field_map = {
        "villain": "villainResponse",
        "expected": "expectedResponse",
        "boring": "boringResponse",
        "logline": "logline",
    }
    output: dict[str, list[str]] = {}
    for key, field in field_map.items():
        output[key] = [tidy_text(row.get(field)) for row in respondents if tidy_text(row.get(field))]
    return output


def build_analysis_rows(workbook: dict[str, list[list[str]]]) -> dict[str, list[dict[str, str]]]:
    output: dict[str, list[dict[str, str]]] = {}
    for sheet_name, rows in workbook.items():
        if not sheet_name.startswith("raw_"):
            continue
        sheet_output: list[dict[str, str]] = []
        for row in rows:
            padded = row + ["", "", ""]
            sheet_output.append({
                "category": tidy_text(padded[0]),
                "text": tidy_text(padded[1]),
                "scene": tidy_text(padded[2]),
            })
        output[sheet_name] = sheet_output
    return output


def infer_missing_analysis_category(sheet_name: str, text: str) -> str:
    value = tidy_text(text)
    if sheet_name == "raw_無聊拖沓" and value in {
        "無", "無。", "沒有", "都還好", "還好", "沒有特別印象", "沒有特別覺得無聊的段落", "都還ok", "看起來沒有覺得太無聊"
    }:
        return "無／還好"
    if sheet_name == "raw_不合理（出戲）" and value in {
        "無", "沒有", "目前還可以 基本邏輯都還過得去", "同上題"
    }:
        return "無／還好"
    return ""


def backfill_missing_analysis_rows(
    analysis_rows: dict[str, list[dict[str, str]]],
    respondents: list[dict[str, Any]],
) -> dict[str, list[dict[str, str]]]:
    field_by_sheet = {
        "raw_局長陰謀": "villainResponse",
        "raw_期待場面": "expectedResponse",
        "raw_無聊拖沓": "boringResponse",
        "raw_不合理（出戲）": "unreasonableResponse",
        "raw_故事核心": "logline",
    }

    for sheet_name, field_name in field_by_sheet.items():
        rows = analysis_rows.get(sheet_name)
        if not rows:
            continue
        header = rows[0]
        body = rows[1:]
        expected_missing = max(len(respondents) - len(body), 0)
        if expected_missing == 0:
            continue
        survey_texts = [tidy_text(row.get(field_name)) for row in respondents if tidy_text(row.get(field_name))]
        survey_counter = Counter(survey_texts)
        analysis_counter = Counter(tidy_text(row.get("text")) for row in body if tidy_text(row.get("text")))
        missing_counter = survey_counter - analysis_counter
        if not missing_counter:
            continue
        candidates: list[str] = []
        prioritized: list[str] = []
        fallback: list[str] = []
        for text, count in missing_counter.items():
            for _ in range(count):
                if infer_missing_analysis_category(sheet_name, text):
                    prioritized.append(text)
                else:
                    fallback.append(text)
        candidates.extend(prioritized)
        candidates.extend(fallback)
        for text in candidates[:expected_missing]:
            body.append({
                "category": infer_missing_analysis_category(sheet_name, text),
                "text": text,
                "scene": "",
            })
        analysis_rows[sheet_name] = [header] + body
    return analysis_rows


def load_existing_dashboard_exports(path: Path) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    if not path.exists():
        return ({}, {}, {})
    text = path.read_text(encoding="utf-8")
    markers = [
        "window.BLIND_TEST_DATA = ",
        "window.BLIND_TEST_REASON_DATA = ",
        "window.BLIND_TEST_OPEN_RESPONSE_DATA = ",
    ]
    try:
        start_data = text.index(markers[0]) + len(markers[0])
        start_reason = text.index(markers[1]) + len(markers[1])
        start_open = text.index(markers[2]) + len(markers[2])
    except ValueError:
        return ({}, {}, {})

    data = json.loads(text[start_data:text.index(markers[1])].rstrip(" ;\n"))
    reason = json.loads(text[start_reason:text.index(markers[2])].rstrip(" ;\n"))
    open_data = json.loads(text[start_open:].rstrip(" ;\n"))
    return data, reason, open_data


def ensure_base_dashboard_data(data: dict[str, Any]) -> dict[str, Any]:
    output = dict(data)
    output["meta"] = output.get("meta") or {"sections": DEFAULT_SECTIONS}
    output["palette"] = output.get("palette") or DEFAULT_PALETTE
    output.setdefault("audience", {})
    output.setdefault("characters", {})
    output.setdefault("worldview", {})
    output.setdefault("plot", {})
    output.setdefault("reception", {})
    output.setdefault("qualitative", {})
    output["characters"].setdefault("favoriteQuotes", {})
    output["characters"].setdefault("dislikeQuotes", {})
    output["reception"].setdefault("associations", {})
    output["reception"]["associations"].setdefault("zones", DEFAULT_ASSOCIATION_ZONES)
    output["qualitative"].setdefault("positive", {"quotes": []})
    output["qualitative"].setdefault("issues", [])
    return output


def to_js_assignment(name: str, payload: Any) -> str:
    return f"window.{name} = {json.dumps(payload, ensure_ascii=False, indent=2)};\n"


def write_outputs(
    dashboard_data: dict[str, Any],
    reason_data: dict[str, Any],
    open_response_data: dict[str, Any],
    workbook_rows: list[dict[str, str]],
    analysis_rows: dict[str, list[dict[str, str]]],
    outputs: dict[str, Path],
) -> None:
    outputs["dashboard_js"].write_text(
        "".join(
            [
                to_js_assignment("BLIND_TEST_DATA", dashboard_data),
                to_js_assignment("BLIND_TEST_REASON_DATA", reason_data),
                to_js_assignment("BLIND_TEST_OPEN_RESPONSE_DATA", open_response_data),
            ]
        ),
        encoding="utf-8",
    )
    outputs["open_response_js"].write_text(
        to_js_assignment("BLIND_TEST_OPEN_RESPONSE_DATA", open_response_data),
        encoding="utf-8",
    )
    outputs["workbook_rows_json"].write_text(
        json.dumps(workbook_rows, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    outputs["analysis_js"].write_text(
        to_js_assignment("WORKBOOK_OPEN_RESPONSE_ROWS", analysis_rows),
        encoding="utf-8",
    )


def verify_alignment(respondents: list[dict[str, Any]], analysis_rows: dict[str, list[dict[str, str]]]) -> list[str]:
    warnings: list[str] = []
    respondent_count = len(respondents)
    for sheet_name, rows in sorted(analysis_rows.items()):
        data_row_count = max(len(rows) - 1, 0)
        if data_row_count != respondent_count:
            warnings.append(
                f"{sheet_name}: analysis rows={data_row_count}, respondents={respondent_count}. "
                "If this is intentional, qualitative cards may lag behind the latest survey rows."
            )
    return warnings


def main() -> int:
    args = parse_args()
    root = args.root.resolve()

    survey_path = (args.survey or find_latest_xlsx(root, ["*劇本盲測問卷*.xlsx", "*盲測問卷*.xlsx"])).resolve()
    analysis_path = (args.analysis or find_latest_xlsx(root, ["*劇本盲測質化分析*.xlsx", "*質化分析*.xlsx"])).resolve()

    survey_workbook = parse_xlsx(survey_path)
    survey_sheet = choose_survey_sheet(survey_workbook)
    workbook_rows = rows_from_sheet(survey_sheet)
    respondents = [map_survey_row(row) for row in workbook_rows if tidy_text(row.get("姓名"))]

    existing_data, _, _ = load_existing_dashboard_exports(DEFAULT_OUTPUTS["dashboard_js"])
    dashboard_data = ensure_base_dashboard_data(existing_data)
    dashboard_data["respondents"] = respondents
    dashboard_data["audience"]["recentMovies"] = derive_recent_movie_groups(respondents)

    reason_data = build_reason_data(respondents)
    open_response_data = build_open_response_data(respondents)

    analysis_workbook = parse_xlsx(analysis_path)
    analysis_rows = build_analysis_rows(analysis_workbook)
    analysis_rows = backfill_missing_analysis_rows(analysis_rows, respondents)

    write_outputs(
        dashboard_data=dashboard_data,
        reason_data=reason_data,
        open_response_data=open_response_data,
        workbook_rows=workbook_rows,
        analysis_rows=analysis_rows,
        outputs=DEFAULT_OUTPUTS,
    )

    warnings = verify_alignment(respondents, analysis_rows)
    print(f"Survey workbook:   {survey_path.name}")
    print(f"Analysis workbook: {analysis_path.name}")
    print(f"Respondents:       {len(respondents)}")
    print(f"Open responses:    villain={len(open_response_data['villain'])}, expected={len(open_response_data['expected'])}, boring={len(open_response_data['boring'])}, logline={len(open_response_data['logline'])}")
    print("Updated files:")
    for path in DEFAULT_OUTPUTS.values():
        print(f"- {path}")
    if warnings:
        print("\nWarnings:")
        for warning in warnings:
            print(f"- {warning}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
