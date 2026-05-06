import re
import time
from collections import defaultdict
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import streamlit as st


# =========================================================
# Streamlit 設定
# =========================================================
st.set_page_config(
    page_title="學代會預算審查系統 Prototype",
    page_icon="📊",
    layout="wide",
)


# =========================================================
# 全域設定
# =========================================================
VALID_DEPARTMENTS = ["行政部門", "立法部門", "司法部門"]
REVIEW_DEPARTMENTS = ["行政部門", "立法部門"]
EXCLUDED_REVIEW_UNITS = {"選舉罷免執行委員會"}
SPLIT_GROUP_UNITS = {"學術部", "文化部"}
PERSONNEL_ACCOUNT_NAMES = {"工作費", "臨時工資", "稿費", "製圖費"}
PROGRAM_GROUP_ALIASES = {
    "永續執行組": "永續組",
}


# =========================================================
# 基本工具函式
# =========================================================
def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\n", " ").strip()


def parse_number(value) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    for token in [",", "元", "%", "％"]:
        text = text.replace(token, "")

    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_rate(value) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        number = float(value)
        return number / 100 if number > 1 else number

    text = str(value).strip()
    if not text:
        return 0.0

    has_percent = "%" in text or "％" in text
    number = parse_number(text)
    return number / 100 if has_percent or number > 1 else number


def format_amount(value) -> str:
    try:
        return f"{float(value):,.0f}"
    except Exception:
        return "0"


def format_rate(value) -> str:
    if value is None or pd.isna(value):
        return "無法計算"
    return f"{float(value) * 100:.2f}%"


def safe_growth_rate(current: float, previous: float) -> Optional[float]:
    if previous == 0:
        if current > 0:
            return None
        return 0.0
    return current / previous - 1


def growth_rate_label(value) -> str:
    if value is None:
        return "上期為 0，無法計算"
    return f"{value * 100:.2f}%"


def pick_sheet(sheet_names: List[str], keywords: Iterable[str], fallback_index: int = 0) -> str:
    for sheet in sheet_names:
        if any(keyword in sheet for keyword in keywords):
            return sheet
    return sheet_names[fallback_index]


def normalize_group_name(name: str) -> str:
    return PROGRAM_GROUP_ALIASES.get(clean_text(name), clean_text(name))


def reset_file_pointer(uploaded_file):
    """Streamlit UploadedFile 被重複讀取前，盡量把游標歸零。"""
    try:
        uploaded_file.seek(0)
    except Exception:
        pass


def infer_period_label(uploaded_file, fallback: str = "未標示年度") -> str:
    """從檔名推測年度／期別，例如 38-1、38-2、37-2。"""
    name = getattr(uploaded_file, "name", "") or ""

    match = re.search(r"(\d{1,3}[-_－]\d{1,2})", name)
    if match:
        return match.group(1).replace("_", "-").replace("－", "-")

    match = re.search(r"(\d{2,3})\s*屆\s*(\d{1,2})", name)
    if match:
        return f"{match.group(1)}-{match.group(2)}"

    return fallback


# =========================================================
# 表格定位與解析
# =========================================================
def find_table_start(raw_df: pd.DataFrame, must_have: Iterable[str], max_rows: int = 80) -> int:
    """找出雙層表頭下一列，也就是真正資料開始列。"""
    must_have = list(must_have)
    for i in range(min(max_rows, len(raw_df))):
        row_text = " ".join(clean_text(x) for x in raw_df.iloc[i].tolist())
        if all(keyword in row_text for keyword in must_have):
            return i + 1
    return 0


def detect_sheet_type(sheet_name: str) -> str:
    if "期入" in sheet_name or "3-1" in sheet_name:
        return "income"
    if "期出" in sheet_name or "3-2" in sheet_name:
        return "expense"
    return "unknown"


def _empty_budget_record(sheet_name: str, row_idx: int, level: str) -> Dict:
    return {
        "source_sheet": sheet_name,
        "budget_type": detect_sheet_type(sheet_name),
        "excel_row": row_idx + 1,
        "level": level,
        "款": "",
        "項": "",
        "目": "",
        "節": "",
        "department": "",
        "unit": "",
        "program_group": "",
        "program": "",
        "account_code": "",
        "account_name": "",
        "amount": 0.0,
        "percentage": 0.0,
        "description": "",
        "review_comment": "",
        "path": "",
    }


def parse_budget_sheet(uploaded_file, sheet_name: str) -> pd.DataFrame:
    """解析預算附屬表。預期欄位：款、項、目、節、代碼、名稱、預算數、百分比、說明、審議結果。"""
    raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    start_row = find_table_start(raw_df, ["款", "項", "目", "節", "代碼", "名稱"])

    records: List[Dict] = []
    current = {
        "department": "",
        "unit": "",
        "program_group": "",
        "program": "",
        "款": "",
        "項": "",
        "目": "",
        "節": "",
    }
    last_idx: Optional[int] = None

    for row_idx in range(start_row, len(raw_df)):
        row = raw_df.iloc[row_idx]
        values = [clean_text(row.iloc[i]) if len(row) > i else "" for i in range(10)]
        kuan, xiang, mu, jie, code, name = values[:6]
        amount = parse_number(row.iloc[6]) if len(row) > 6 else 0.0
        percentage = parse_rate(row.iloc[7]) if len(row) > 7 else 0.0
        description = clean_text(row.iloc[8]) if len(row) > 8 else ""
        review_comment = clean_text(row.iloc[9]) if len(row) > 9 else ""

        if not any([kuan, xiang, mu, jie, code, name, amount, percentage, description, review_comment]):
            continue

        # 合計列
        if not code and name == "合計":
            rec = _empty_budget_record(sheet_name, row_idx, "合計")
            rec.update({
                "款": kuan,
                "項": xiang,
                "目": mu,
                "節": jie,
                "account_name": "合計",
                "amount": amount,
                "percentage": percentage,
                "description": description,
                "review_comment": review_comment,
                "path": "合計",
            })
            records.append(rec)
            last_idx = len(records) - 1
            continue

        # 階層列
        if not code and name:
            if name in VALID_DEPARTMENTS:
                level = "部門"
                current.update({
                    "department": name,
                    "unit": "",
                    "program_group": "",
                    "program": "",
                    "款": kuan,
                    "項": "",
                    "目": "",
                    "節": "",
                })
            elif xiang:
                level = "單位"
                current.update({"unit": name, "program_group": "", "program": "", "項": xiang, "目": "", "節": ""})
            elif mu and jie and jie not in {"0", "0.0"}:
                level = "計畫"
                current.update({"program": name, "目": mu, "節": jie})
            elif mu:
                level = "計畫群"
                current.update({"program_group": name, "program": "", "目": mu, "節": ""})
            elif jie:
                level = "計畫"
                current.update({"program": name, "節": jie})
            else:
                level = "計畫"
                current.update({"program": name})

            path = " > ".join(x for x in [current["department"], current["unit"], current["program_group"], current["program"]] if x)
            rec = _empty_budget_record(sheet_name, row_idx, level)
            rec.update({
                "款": current["款"],
                "項": current["項"],
                "目": current["目"],
                "節": current["節"],
                "department": current["department"],
                "unit": current["unit"],
                "program_group": current["program_group"],
                "program": current["program"],
                "account_name": name,
                "amount": amount,
                "percentage": percentage,
                "description": description,
                "review_comment": review_comment,
                "path": path,
            })
            records.append(rec)
            last_idx = len(records) - 1
            continue

        # 科目列
        if code or name:
            path = " > ".join(
                x for x in [
                    current["department"], current["unit"], current["program_group"], current["program"], f"{code} {name}".strip()
                ] if x
            )
            rec = _empty_budget_record(sheet_name, row_idx, "科目")
            rec.update({
                "款": current["款"],
                "項": current["項"],
                "目": current["目"],
                "節": current["節"],
                "department": current["department"],
                "unit": current["unit"],
                "program_group": current["program_group"],
                "program": current["program"],
                "account_code": code,
                "account_name": name,
                "amount": amount,
                "percentage": percentage,
                "description": description,
                "review_comment": review_comment,
                "path": path,
            })
            records.append(rec)
            last_idx = len(records) - 1
            continue

        # 延伸說明列
        if last_idx is not None and (description or review_comment):
            if description:
                old = records[last_idx].get("description", "")
                records[last_idx]["description"] = f"{old}\n{description}" if old else description
            if review_comment:
                old = records[last_idx].get("review_comment", "")
                records[last_idx]["review_comment"] = f"{old}\n{review_comment}" if old else review_comment

    result = pd.DataFrame(records)
    if result.empty:
        return result

    result["amount"] = pd.to_numeric(result["amount"], errors="coerce").fillna(0)
    result["percentage"] = pd.to_numeric(result["percentage"], errors="coerce").fillna(0)
    result = add_search_text(result)
    return result


def parse_final_sheet(uploaded_file, sheet_name: str) -> pd.DataFrame:
    """解析決算附屬表。預期欄位：款、項、目、節、代碼、名稱、決算數、決算百分比、預算數、預算百分比、餘絀數、執行率、說明、審議結果、備註。"""
    raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    start_row = find_table_start(raw_df, ["款", "項", "目", "節", "決算", "預算"])

    records: List[Dict] = []
    current = {
        "department": "",
        "unit": "",
        "program_group": "",
        "program": "",
        "款": "",
        "項": "",
        "目": "",
        "節": "",
    }
    last_idx: Optional[int] = None

    def empty_final_record(row_idx: int, level: str) -> Dict:
        return {
            "source_sheet": sheet_name,
            "budget_type": detect_sheet_type(sheet_name),
            "excel_row": row_idx + 1,
            "level": level,
            "款": "",
            "項": "",
            "目": "",
            "節": "",
            "department": "",
            "unit": "",
            "program_group": "",
            "program": "",
            "account_code": "",
            "account_name": "",
            "final_amount": 0.0,
            "final_percentage": 0.0,
            "budget_amount": 0.0,
            "budget_percentage": 0.0,
            "balance_amount": 0.0,
            "execution_rate": 0.0,
            "description": "",
            "review_comment": "",
            "department_note": "",
            "path": "",
        }

    for row_idx in range(start_row, len(raw_df)):
        row = raw_df.iloc[row_idx]
        kuan = clean_text(row.iloc[0]) if len(row) > 0 else ""
        xiang = clean_text(row.iloc[1]) if len(row) > 1 else ""
        mu = clean_text(row.iloc[2]) if len(row) > 2 else ""
        jie = clean_text(row.iloc[3]) if len(row) > 3 else ""
        code = clean_text(row.iloc[4]) if len(row) > 4 else ""
        name = clean_text(row.iloc[5]) if len(row) > 5 else ""
        final_amount = parse_number(row.iloc[6]) if len(row) > 6 else 0.0
        final_percentage = parse_rate(row.iloc[7]) if len(row) > 7 else 0.0
        budget_amount = parse_number(row.iloc[8]) if len(row) > 8 else 0.0
        budget_percentage = parse_rate(row.iloc[9]) if len(row) > 9 else 0.0
        balance_amount = parse_number(row.iloc[10]) if len(row) > 10 else 0.0
        execution_rate = parse_rate(row.iloc[11]) if len(row) > 11 else 0.0
        description = clean_text(row.iloc[12]) if len(row) > 12 else ""
        review_comment = clean_text(row.iloc[13]) if len(row) > 13 else ""
        department_note = clean_text(row.iloc[14]) if len(row) > 14 else ""

        if execution_rate == 0 and budget_amount > 0:
            execution_rate = final_amount / budget_amount

        if not any([kuan, xiang, mu, jie, code, name, final_amount, budget_amount, balance_amount, description, review_comment, department_note]):
            continue

        if not code and name == "合計":
            rec = empty_final_record(row_idx, "合計")
            rec.update({
                "款": kuan,
                "項": xiang,
                "目": mu,
                "節": jie,
                "account_name": "合計",
                "final_amount": final_amount,
                "final_percentage": final_percentage,
                "budget_amount": budget_amount,
                "budget_percentage": budget_percentage,
                "balance_amount": balance_amount,
                "execution_rate": execution_rate,
                "description": description,
                "review_comment": review_comment,
                "department_note": department_note,
                "path": "合計",
            })
            records.append(rec)
            last_idx = len(records) - 1
            continue

        if not code and name:
            if name in VALID_DEPARTMENTS:
                level = "部門"
                current.update({"department": name, "unit": "", "program_group": "", "program": "", "款": kuan, "項": "", "目": "", "節": ""})
            elif xiang:
                level = "單位"
                current.update({"unit": name, "program_group": "", "program": "", "項": xiang, "目": "", "節": ""})
            elif mu and jie and jie not in {"0", "0.0"}:
                level = "計畫"
                current.update({"program": name, "目": mu, "節": jie})
            elif mu:
                level = "計畫群"
                current.update({"program_group": name, "program": "", "目": mu, "節": ""})
            elif jie:
                level = "計畫"
                current.update({"program": name, "節": jie})
            else:
                level = "計畫"
                current.update({"program": name})

            path = " > ".join(x for x in [current["department"], current["unit"], current["program_group"], current["program"]] if x)
            rec = empty_final_record(row_idx, level)
            rec.update({
                "款": current["款"],
                "項": current["項"],
                "目": current["目"],
                "節": current["節"],
                "department": current["department"],
                "unit": current["unit"],
                "program_group": current["program_group"],
                "program": current["program"],
                "account_name": name,
                "final_amount": final_amount,
                "final_percentage": final_percentage,
                "budget_amount": budget_amount,
                "budget_percentage": budget_percentage,
                "balance_amount": balance_amount,
                "execution_rate": execution_rate,
                "description": description,
                "review_comment": review_comment,
                "department_note": department_note,
                "path": path,
            })
            records.append(rec)
            last_idx = len(records) - 1
            continue

        if code or name:
            path = " > ".join(
                x for x in [current["department"], current["unit"], current["program_group"], current["program"], f"{code} {name}".strip()] if x
            )
            rec = empty_final_record(row_idx, "科目")
            rec.update({
                "款": current["款"],
                "項": current["項"],
                "目": current["目"],
                "節": current["節"],
                "department": current["department"],
                "unit": current["unit"],
                "program_group": current["program_group"],
                "program": current["program"],
                "account_code": code,
                "account_name": name,
                "final_amount": final_amount,
                "final_percentage": final_percentage,
                "budget_amount": budget_amount,
                "budget_percentage": budget_percentage,
                "balance_amount": balance_amount,
                "execution_rate": execution_rate,
                "description": description,
                "review_comment": review_comment,
                "department_note": department_note,
                "path": path,
            })
            records.append(rec)
            last_idx = len(records) - 1
            continue

        if last_idx is not None and (description or review_comment or department_note):
            if description:
                old = records[last_idx].get("description", "")
                records[last_idx]["description"] = f"{old}\n{description}" if old else description
            if review_comment:
                old = records[last_idx].get("review_comment", "")
                records[last_idx]["review_comment"] = f"{old}\n{review_comment}" if old else review_comment
            if department_note:
                old = records[last_idx].get("department_note", "")
                records[last_idx]["department_note"] = f"{old}\n{department_note}" if old else department_note

    result = pd.DataFrame(records)
    if result.empty:
        return result

    for col in ["final_amount", "final_percentage", "budget_amount", "budget_percentage", "balance_amount", "execution_rate"]:
        result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0)
    result = add_search_text(result)
    return result


def add_search_text(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    searchable_cols = [
        "level", "department", "unit", "program_group", "program", "account_code", "account_name", "description", "review_comment", "path"
    ]
    existing = [c for c in searchable_cols if c in result.columns]
    result["search_text"] = result[existing].fillna("").agg(" ".join, axis=1).str.lower()
    return result


# =========================================================
# 審查單位處理
# =========================================================
def add_review_unit_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    if "program_group" not in result.columns:
        result["program_group"] = ""

    result["normalized_program_group"] = result["program_group"].fillna("").map(normalize_group_name)
    result["review_unit"] = result["unit"].fillna("")
    result["review_unit_type"] = "單位"

    should_split = result["unit"].isin(SPLIT_GROUP_UNITS) & result["normalized_program_group"].ne("")
    result.loc[should_split, "review_unit"] = (
        result.loc[should_split, "unit"].fillna("") + "｜" + result.loc[should_split, "normalized_program_group"].fillna("")
    )
    result.loc[should_split, "review_unit_type"] = "小組"
    return result


def apply_review_scope(df: pd.DataFrame) -> pd.DataFrame:
    result = add_review_unit_columns(df)
    result = result[result["department"].isin(REVIEW_DEPARTMENTS)].copy()
    result = result[~result["unit"].isin(EXCLUDED_REVIEW_UNITS)].copy()
    return result


def get_personnel_df(df: pd.DataFrame) -> pd.DataFrame:
    scoped = apply_review_scope(df)
    return scoped[scoped["account_name"].isin(PERSONNEL_ACCOUNT_NAMES)].copy()


def make_review_base(current_expense_df: pd.DataFrame, previous_final_df: pd.DataFrame) -> pd.DataFrame:
    current_units = apply_review_scope(current_expense_df)[["department", "review_unit", "review_unit_type"]].drop_duplicates()
    previous_units = apply_review_scope(previous_final_df)[["department", "review_unit", "review_unit_type"]].drop_duplicates()
    base = pd.concat([current_units, previous_units], ignore_index=True).drop_duplicates()
    base = base[base["review_unit"].fillna("") != ""]
    return base.sort_values(["department", "review_unit"]).reset_index(drop=True)


# =========================================================
# 審查規則
# =========================================================
def calculate_non_fee_income(current_income_df: pd.DataFrame) -> pd.DataFrame:
    if current_income_df.empty:
        return pd.DataFrame(columns=["department", "review_unit", "review_unit_type", "non_fee_income"])

    income = current_income_df[current_income_df["level"] == "科目"].copy()
    income = apply_review_scope(income)

    income["is_fee_income"] = income["account_name"].str.contains("會費", na=False)

    # 會長一般行政計畫收入屬於會費來源，不計入非會費收入。
    president_general_admin_income = income["unit"].eq("會長") & income["program"].fillna("").str.contains("一般行政計畫", na=False)
    income.loc[president_general_admin_income, "is_fee_income"] = True

    non_fee = income[~income["is_fee_income"]].copy()
    return (
        non_fee.groupby(["department", "review_unit", "review_unit_type"], dropna=False)["amount"]
        .sum()
        .reset_index()
        .rename(columns={"amount": "non_fee_income"})
    )


def rule_1_personnel_cap(current_expense_df: pd.DataFrame, previous_final_df: pd.DataFrame, current_income_df: pd.DataFrame) -> pd.DataFrame:
    """
    新版規則一：人事相關費用合理性審查。

    檢查科目：工作費、臨時工資、稿費、製圖費。
    比較粒度：審查單位／小組。

    上限 A = 上期人事執行決算 × 120%
    上限 B = 上期人事執行決算 × (1 + 整體預算成長率 × 110%)
    合理人事費上限 = max(上限 A, 上限 B) + 本期新增非會費收入
    建議凍結金額 = max(0, 本期人事相關預算 - 合理人事費上限)
    """
    current_scoped = apply_review_scope(current_expense_df)
    previous_scoped = apply_review_scope(previous_final_df)

    current_personnel = current_scoped[current_scoped["account_name"].isin(PERSONNEL_ACCOUNT_NAMES)].copy()
    previous_personnel = previous_scoped[previous_scoped["account_name"].isin(PERSONNEL_ACCOUNT_NAMES)].copy()

    current_personnel_summary = (
        current_personnel.groupby(["department", "review_unit", "review_unit_type"], dropna=False)["amount"]
        .sum()
        .reset_index()
        .rename(columns={"amount": "current_personnel_budget"})
    )
    previous_personnel_summary = (
        previous_personnel.groupby(["department", "review_unit", "review_unit_type"], dropna=False)["final_amount"]
        .sum()
        .reset_index()
        .rename(columns={"final_amount": "previous_personnel_final"})
    )
    current_total_summary = (
        current_scoped.groupby(["department", "review_unit", "review_unit_type"], dropna=False)["amount"]
        .sum()
        .reset_index()
        .rename(columns={"amount": "current_total_budget"})
    )
    previous_total_summary = (
        previous_scoped.groupby(["department", "review_unit", "review_unit_type"], dropna=False)["final_amount"]
        .sum()
        .reset_index()
        .rename(columns={"final_amount": "previous_total_final"})
    )
    non_fee_summary = calculate_non_fee_income(current_income_df)

    result = make_review_base(current_expense_df, previous_final_df)
    for summary in [
        current_personnel_summary,
        previous_personnel_summary,
        current_total_summary,
        previous_total_summary,
        non_fee_summary,
    ]:
        result = result.merge(summary, on=["department", "review_unit", "review_unit_type"], how="left")

    for col in [
        "current_personnel_budget",
        "previous_personnel_final",
        "current_total_budget",
        "previous_total_final",
        "non_fee_income",
    ]:
        result[col] = result[col].fillna(0)

    rows = []
    for _, row in result.iterrows():
        total_growth = safe_growth_rate(row["current_total_budget"], row["previous_total_final"])

        cap_a = row["previous_personnel_final"] * 1.2
        if total_growth is None:
            # 若上期整體決算為 0，無法計算成長率，保守以 A 作為 B。
            cap_b = cap_a
            total_growth_label = "上期整體決算為 0，無法計算"
        else:
            cap_b = row["previous_personnel_final"] * (1 + total_growth * 1.1)
            cap_b = max(0, cap_b)
            total_growth_label = growth_rate_label(total_growth)

        reasonable_cap_before_income = max(cap_a, cap_b)
        reasonable_cap = reasonable_cap_before_income + row["non_fee_income"]
        freeze = max(0, row["current_personnel_budget"] - reasonable_cap)
        violate = freeze > 0

        reason = (
            f"上限A為上期人事決算120%：{format_amount(cap_a)}元；"
            f"整體預算成長率為{total_growth_label}，上限B為{format_amount(cap_b)}元；"
            f"本期新增非會費收入為{format_amount(row['non_fee_income'])}元；"
            f"合理人事費上限為{format_amount(reasonable_cap)}元。"
        )
        if not violate:
            reason += "未超過合理上限。"

        new_row = row.to_dict()
        new_row.update({
            "total_growth_rate": total_growth,
            "cap_a_previous_120": cap_a,
            "cap_b_growth_adjusted": cap_b,
            "reasonable_cap_before_income": reasonable_cap_before_income,
            "reasonable_personnel_cap": reasonable_cap,
            "suggested_freeze_amount": freeze,
            "violate_rule": violate,
            "risk_reason": reason,
        })
        rows.append(new_row)

    return pd.DataFrame(rows).sort_values(
        ["violate_rule", "suggested_freeze_amount", "department", "review_unit"],
        ascending=[False, False, True, True],
    )


def rule_2_personnel_growth(current_expense_df: pd.DataFrame, previous_final_df: pd.DataFrame, current_income_df: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    current_scoped = apply_review_scope(current_expense_df)
    previous_scoped = apply_review_scope(previous_final_df)
    current_personnel = current_scoped[current_scoped["account_name"].isin(PERSONNEL_ACCOUNT_NAMES)].copy()
    previous_personnel = previous_scoped[previous_scoped["account_name"].isin(PERSONNEL_ACCOUNT_NAMES)].copy()

    current_personnel_summary = (
        current_personnel.groupby(["department", "review_unit", "review_unit_type"], dropna=False)["amount"]
        .sum()
        .reset_index()
        .rename(columns={"amount": "current_personnel_budget"})
    )
    previous_personnel_summary = (
        previous_personnel.groupby(["department", "review_unit", "review_unit_type"], dropna=False)["final_amount"]
        .sum()
        .reset_index()
        .rename(columns={"final_amount": "previous_personnel_final"})
    )
    current_total_summary = (
        current_scoped.groupby(["department", "review_unit", "review_unit_type"], dropna=False)["amount"]
        .sum()
        .reset_index()
        .rename(columns={"amount": "current_total_budget"})
    )
    previous_total_summary = (
        previous_scoped.groupby(["department", "review_unit", "review_unit_type"], dropna=False)["final_amount"]
        .sum()
        .reset_index()
        .rename(columns={"final_amount": "previous_total_final"})
    )

    non_fee_summary = calculate_non_fee_income(current_income_df) if current_income_df is not None else pd.DataFrame(columns=["department", "review_unit", "review_unit_type", "non_fee_income"])

    result = make_review_base(current_expense_df, previous_final_df)
    for summary in [current_personnel_summary, previous_personnel_summary, current_total_summary, previous_total_summary, non_fee_summary]:
        result = result.merge(summary, on=["department", "review_unit", "review_unit_type"], how="left")

    for col in ["current_personnel_budget", "previous_personnel_final", "current_total_budget", "previous_total_final", "non_fee_income"]:
        result[col] = result[col].fillna(0)

    rows = []
    for _, row in result.iterrows():
        personnel_growth = safe_growth_rate(row["current_personnel_budget"], row["previous_personnel_final"])
        total_growth = safe_growth_rate(row["current_total_budget"], row["previous_total_final"])

        if personnel_growth is None or total_growth is None:
            violate = row["current_personnel_budget"] > 0 and row["previous_personnel_final"] == 0
            allowed = 0 if violate else row["current_personnel_budget"]
            freeze = row["current_personnel_budget"] if violate else 0
            reason = "上期人事費或整體決算為 0，需人工檢查；暫以本期新增人事費作為建議凍結額"
        else:
            threshold_growth = total_growth * 1.1
            base_allowed = max(0, row["previous_personnel_final"] * (1 + threshold_growth))
            # 規則二較彈性的版本：若該審查單位／小組能找到非會費收入，
            # 可用非會費收入抵減人事費成長過高所產生的凍結建議。
            allowed = base_allowed + row["non_fee_income"]
            violate = row["current_personnel_budget"] > allowed
            freeze = max(0, row["current_personnel_budget"] - allowed)
            reason = (
                f"人事費成長率 {growth_rate_label(personnel_growth)}，"
                f"整體預算成長率 {growth_rate_label(total_growth)}，"
                f"基本門檻為整體成長率的 110%；"
                f"基本人事費上限 {format_amount(base_allowed)} 元，"
                f"可抵減非會費收入 {format_amount(row['non_fee_income'])} 元，"
                f"調整後上限 {format_amount(allowed)} 元"
            )

        new_row = row.to_dict()
        new_row.update({
            "personnel_growth_rate": personnel_growth,
            "total_growth_rate": total_growth,
            "non_fee_income": row["non_fee_income"],
            "allowed_personnel_budget": allowed,
            "suggested_freeze_amount": freeze,
            "violate_rule": violate,
            "risk_reason": reason,
        })
        rows.append(new_row)

    return pd.DataFrame(rows).sort_values(["violate_rule", "suggested_freeze_amount", "department", "review_unit"], ascending=[False, False, True, True])


def rule_3_low_execution_freeze(current_expense_df: pd.DataFrame, previous_final_df: pd.DataFrame) -> pd.DataFrame:
    current = apply_review_scope(current_expense_df)
    previous = apply_review_scope(previous_final_df)
    current = current[current["account_name"].isin(PERSONNEL_ACCOUNT_NAMES)].copy()
    previous = previous[previous["account_name"].isin(PERSONNEL_ACCOUNT_NAMES)].copy()

    group_cols = ["department", "review_unit", "review_unit_type", "account_code", "account_name"]
    current_group = (
        current.groupby(group_cols, dropna=False)
        .agg(
            current_budget_amount=("amount", "sum"),
            current_programs=("program", lambda x: "、".join(sorted(set(clean_text(v) for v in x if clean_text(v))))),
            current_paths=("path", lambda x: "；".join(sorted(set(clean_text(v) for v in x if clean_text(v))))),
        )
        .reset_index()
    )
    previous_group = (
        previous.groupby(group_cols, dropna=False)
        .agg(
            previous_final_amount=("final_amount", "sum"),
            previous_budget_amount=("budget_amount", "sum"),
            previous_paths=("path", lambda x: "；".join(sorted(set(clean_text(v) for v in x if clean_text(v))))),
        )
        .reset_index()
    )

    merged = current_group.merge(previous_group, on=group_cols, how="left")
    for col in ["previous_final_amount", "previous_budget_amount"]:
        merged[col] = merged[col].fillna(0)

    merged["previous_execution_rate"] = merged.apply(
        lambda row: row["previous_final_amount"] / row["previous_budget_amount"] if row["previous_budget_amount"] > 0 else 0,
        axis=1,
    )
    merged["violate_rule"] = (merged["previous_execution_rate"] < 0.8) & (merged["current_budget_amount"] > merged["previous_final_amount"])
    merged["suggested_freeze_amount"] = (merged["current_budget_amount"] - merged["previous_final_amount"]).clip(lower=0)
    merged["risk_reason"] = merged.apply(
        lambda row: "上期該單位同一人事相關科目合計執行率低於 80%，且本期預算合計高於上期決算合計數" if row["violate_rule"] else "未觸發本規則",
        axis=1,
    )
    return merged[merged["violate_rule"]].sort_values("suggested_freeze_amount", ascending=False)


def build_review_suggestions(current_expense_df: pd.DataFrame, previous_final_df: pd.DataFrame, current_income_df: pd.DataFrame) -> pd.DataFrame:
    result = add_review_unit_columns(current_expense_df).copy()
    result["system_review_suggestion"] = ""

    if previous_final_df.empty:
        result["review_result_suggestion"] = result["review_comment"].fillna("")
        return result

    rule1 = rule_1_personnel_cap(current_expense_df, previous_final_df, current_income_df)
    rule3 = rule_3_low_execution_freeze(current_expense_df, previous_final_df)

    unit_suggestions = defaultdict(list)
    for _, row in rule1[rule1["violate_rule"]].iterrows():
        unit_suggestions[(row["department"], row["review_unit"])].append(
            f"規則一建議凍結 {format_amount(row['suggested_freeze_amount'])} 元：{row['risk_reason']}"
        )
    account_suggestions = defaultdict(list)
    if not rule3.empty:
        for _, row in rule3.iterrows():
            account_suggestions[(row["department"], row["review_unit"], row["account_code"], row["account_name"])].append(
                f"規則三建議凍結 {format_amount(row['suggested_freeze_amount'])} 元：{row['risk_reason']}"
            )

    def make_suggestion(row) -> str:
        if row["account_name"] not in PERSONNEL_ACCOUNT_NAMES:
            return ""
        parts = []
        parts.extend(unit_suggestions.get((row["department"], row["review_unit"]), []))
        parts.extend(account_suggestions.get((row["department"], row["review_unit"], row["account_code"], row["account_name"]), []))
        return "\n".join(parts)

    result["system_review_suggestion"] = result.apply(make_suggestion, axis=1)
    result["review_result_suggestion"] = result.apply(
        lambda row: "\n".join([x for x in [clean_text(row.get("review_comment", "")), clean_text(row.get("system_review_suggestion", ""))] if x]),
        axis=1,
    )
    return result


# =========================================================
# 跨年度比較資料整理
# =========================================================
def build_multi_year_budget_df(budget_files: List) -> pd.DataFrame:
    """
    讀取多個年度／期別的預算表，整理成跨年度比較資料。

    粒度：period_label × department × review_unit × account_code × account_name。
    """
    all_rows = []

    for idx, file in enumerate(budget_files):
        if file is None:
            continue

        reset_file_pointer(file)

        try:
            excel = pd.ExcelFile(file)
            sheet_names = excel.sheet_names
            expense_sheet = pick_sheet(sheet_names, ["期出", "3-2"])
            period_label = infer_period_label(file, fallback=f"第{idx + 1}份")
        except Exception:
            continue

        try:
            reset_file_pointer(file)
            parsed = parse_budget_sheet(file, expense_sheet)
        except Exception:
            continue

        accounts = parsed[parsed["level"] == "科目"].copy()
        accounts = accounts[accounts["department"].isin(VALID_DEPARTMENTS)].copy()
        accounts = add_review_unit_columns(accounts)

        accounts["period_label"] = period_label
        accounts["source_file"] = getattr(file, "name", f"第{idx + 1}份")
        accounts["account_label"] = accounts["account_code"].fillna("") + "｜" + accounts["account_name"].fillna("")

        all_rows.append(accounts)

    if not all_rows:
        return pd.DataFrame()

    result = pd.concat(all_rows, ignore_index=True)
    result = add_search_text(result)
    return result


# =========================================================
# 搜尋與效能比較
# =========================================================
def linear_search(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    keyword = keyword.lower().strip()
    if not keyword:
        return df
    return df[df["search_text"].str.contains(re.escape(keyword), na=False)]


def build_inverted_index(df: pd.DataFrame) -> dict:
    index = defaultdict(set)
    for i, text in enumerate(df["search_text"].fillna("").tolist()):
        tokens = re.findall(r"[\w\u4e00-\u9fff]+", text.lower())
        for token in tokens:
            index[token].add(i)
    return index


def hash_search(df: pd.DataFrame, index: dict, keyword: str) -> pd.DataFrame:
    keyword = keyword.lower().strip()
    if not keyword:
        return df
    if keyword in index:
        return df.iloc[sorted(index[keyword])]
    matched = set()
    for token, indices in index.items():
        if keyword in token:
            matched.update(indices)
    return df.iloc[sorted(matched)]


def measure_search_performance(df: pd.DataFrame, keyword: str, repeat: int = 100) -> Optional[Dict]:
    if df.empty or not keyword.strip():
        return None

    start = time.perf_counter()
    for _ in range(repeat):
        linear_result = linear_search(df, keyword)
    linear_time = time.perf_counter() - start

    start = time.perf_counter()
    index = build_inverted_index(df)
    build_time = time.perf_counter() - start

    start = time.perf_counter()
    for _ in range(repeat):
        hash_result = hash_search(df, index, keyword)
    hash_time = time.perf_counter() - start

    return {
        "repeat": repeat,
        "linear_time": linear_time,
        "index_build_time": build_time,
        "hash_time": hash_time,
        "linear_result_count": len(linear_result),
        "hash_result_count": len(hash_result),
    }


# =========================================================
# 顯示輔助函式
# =========================================================
def display_df(df: pd.DataFrame, columns: Dict[str, str], amount_cols: Optional[List[str]] = None, rate_cols: Optional[List[str]] = None):
    amount_cols = amount_cols or []
    rate_cols = rate_cols or []
    temp = df.copy()
    for col in amount_cols:
        if col in temp.columns:
            temp[col] = temp[col].apply(format_amount)
    for col in rate_cols:
        if col in temp.columns:
            temp[col] = temp[col].apply(format_rate)

    existing = [col for col in columns if col in temp.columns]
    st.dataframe(temp[existing].rename(columns={c: columns[c] for c in existing}), use_container_width=True, hide_index=True)


def render_budget_tree(df: pd.DataFrame):
    subject_df = df[df["level"] == "科目"].copy()
    if subject_df.empty:
        st.warning("沒有可顯示的科目資料。")
        return

    for dept in [x for x in subject_df["department"].dropna().unique() if x]:
        dept_df = subject_df[subject_df["department"] == dept]
        st.markdown(f"## {dept}｜科目合計 {format_amount(dept_df['amount'].sum())} 元")

        for unit in [x for x in dept_df["unit"].dropna().unique() if x]:
            unit_df = dept_df[dept_df["unit"] == unit]
            with st.expander(f"{unit}｜{format_amount(unit_df['amount'].sum())} 元", expanded=False):
                programs = [x for x in unit_df["program"].dropna().unique() if x]
                if not programs:
                    display_df(
                        unit_df,
                        {
                            "account_code": "代碼", "account_name": "科目", "amount": "金額", "percentage": "百分比",
                            "description": "說明", "review_result_suggestion": "審查結果之建議",
                        },
                        amount_cols=["amount"],
                        rate_cols=["percentage"],
                    )
                    continue

                for program in programs:
                    program_df = unit_df[unit_df["program"] == program]
                    st.markdown(f"#### {program}｜{format_amount(program_df['amount'].sum())} 元")
                    display_df(
                        program_df,
                        {
                            "account_code": "代碼", "account_name": "科目", "amount": "金額", "percentage": "百分比",
                            "description": "說明", "review_result_suggestion": "審查結果之建議",
                        },
                        amount_cols=["amount"],
                        rate_cols=["percentage"],
                    )


# =========================================================
# 頁面與資料載入
# =========================================================
st.title("📊 學代會歷年預算審查系統 Prototype")
st.caption("本期預算 + 上期決算｜階層瀏覽｜科目瀏覽｜三項審查規則｜搜尋效能比較")

with st.sidebar:
    st.header("資料來源")
    current_budget_file = st.file_uploader("上傳本期預算 Excel 檔", type=["xlsx", "xls"], key="current_budget")
    previous_final_file = st.file_uploader("上傳上期決算 Excel 檔", type=["xlsx", "xls"], key="previous_final")
    multi_year_budget_files = st.file_uploader(
        "上傳多年度預算 Excel 檔（可複選，用於跨年度比較）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="multi_year_budget",
    )

    st.divider()
    st.markdown(
        """
        **審查設定**
        - 排除司法部門
        - 排除選舉罷免執行委員會
        - 學術部、文化部依小組拆分
        - 永續執行組視為永續組
        """
    )

if current_budget_file is None:
    st.info("請先在左側上傳本期預算 Excel 檔。")
    st.stop()

try:
    current_excel = pd.ExcelFile(current_budget_file)
    current_sheets = current_excel.sheet_names
except Exception as e:
    st.error(f"讀取本期預算 Excel 失敗：{e}")
    st.stop()

current_expense_sheet = st.sidebar.selectbox(
    "本期支出預算工作表",
    current_sheets,
    index=current_sheets.index(pick_sheet(current_sheets, ["期出", "3-2"])),
)
current_income_sheet = st.sidebar.selectbox(
    "本期收入預算工作表",
    current_sheets,
    index=current_sheets.index(pick_sheet(current_sheets, ["期入", "3-1"])),
)

try:
    current_expense_df = parse_budget_sheet(current_budget_file, current_expense_sheet)
    current_income_df = parse_budget_sheet(current_budget_file, current_income_sheet)
except Exception as e:
    st.error(f"本期預算資料清理失敗：{e}")
    st.stop()

if current_expense_df.empty:
    st.warning("本期支出預算工作表沒有成功解析出資料。")
    st.stop()

current_expense_accounts_raw = current_expense_df[current_expense_df["level"] == "科目"].copy()
current_expense_accounts_raw = current_expense_accounts_raw[current_expense_accounts_raw["department"].isin(VALID_DEPARTMENTS)].copy()

previous_final_df = pd.DataFrame()
previous_final_accounts = pd.DataFrame()
if previous_final_file is not None:
    try:
        previous_excel = pd.ExcelFile(previous_final_file)
        previous_sheets = previous_excel.sheet_names
        previous_final_sheet = st.sidebar.selectbox(
            "上期支出決算工作表",
            previous_sheets,
            index=previous_sheets.index(pick_sheet(previous_sheets, ["期出", "3-2"])),
        )
        previous_final_df = parse_final_sheet(previous_final_file, previous_final_sheet)
        previous_final_accounts = previous_final_df[previous_final_df["level"] == "科目"].copy()
        previous_final_accounts = previous_final_accounts[previous_final_accounts["department"].isin(VALID_DEPARTMENTS)].copy()
    except Exception as e:
        st.error(f"上期決算資料清理失敗：{e}")
        st.stop()

# 回填系統審查建議
current_expense_accounts = build_review_suggestions(current_expense_accounts_raw, previous_final_accounts, current_income_df)

current_hierarchy_df = current_expense_df[current_expense_df["level"].isin(["合計", "部門", "單位", "計畫群", "計畫", "科目"])].copy()
current_hierarchy_df = current_hierarchy_df[
    current_hierarchy_df["department"].isin(VALID_DEPARTMENTS) | current_hierarchy_df["level"].eq("合計")
].copy()
current_hierarchy_df = current_hierarchy_df.merge(
    current_expense_accounts[["excel_row", "review_result_suggestion", "system_review_suggestion"]],
    on="excel_row",
    how="left",
)
current_hierarchy_df["review_result_suggestion"] = current_hierarchy_df["review_result_suggestion"].fillna(
    current_hierarchy_df["review_comment"].fillna("")
)

# 更新 search_text，納入系統建議
current_expense_accounts = add_search_text(current_expense_accounts)
current_hierarchy_df = add_search_text(current_hierarchy_df)

# 跨年度預算比較資料
multi_year_files_for_compare = []
if multi_year_budget_files:
    multi_year_files_for_compare.extend(multi_year_budget_files)
else:
    multi_year_files_for_compare.append(current_budget_file)

multi_year_budget_df = build_multi_year_budget_df(multi_year_files_for_compare)

# Dashboard
total_amount = current_expense_accounts["amount"].sum()
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("本期科目總金額", f"{total_amount:,.0f} 元")
col2.metric("科目筆數", f"{len(current_expense_accounts):,}")
col3.metric("部門數", f"{current_expense_accounts['department'].nunique():,}")
col4.metric("單位數", f"{current_expense_accounts['unit'].replace('', pd.NA).dropna().nunique():,}")
col5.metric("計畫數", f"{current_expense_accounts['program'].replace('', pd.NA).dropna().nunique():,}")


tab_overview, tab_unit, tab_account, tab_year_compare, tab_rules, tab_search, tab_ranking, tab_performance, tab_raw = st.tabs(
    ["總覽", "單位預算瀏覽", "科目層級瀏覽", "跨年度比較", "審查規則", "搜尋", "高金額排序", "效能比較", "清理後資料"]
)


# =========================================================
# 總覽
# =========================================================
with tab_overview:
    st.subheader("各單位支出總覽")
    unit_summary = (
        current_expense_accounts.groupby(["department", "unit"], dropna=False)["amount"]
        .sum()
        .reset_index()
        .sort_values(["department", "amount"], ascending=[True, False])
    )
    unit_summary = unit_summary[unit_summary["unit"] != ""]
    display_df(unit_summary, {"department": "部門", "unit": "單位", "amount": "支出總額"}, amount_cols=["amount"])
    if not unit_summary.empty:
        chart_df = unit_summary.copy()
        chart_df["label"] = chart_df["department"] + "｜" + chart_df["unit"]
        st.bar_chart(chart_df.set_index("label")["amount"])

    with st.expander("三大部門彙總", expanded=False):
        dept_summary = current_expense_accounts.groupby("department")["amount"].sum().reset_index().sort_values("amount", ascending=False)
        display_df(dept_summary, {"department": "部門", "amount": "支出總額"}, amount_cols=["amount"])


# =========================================================
# 單位預算瀏覽
# =========================================================
with tab_unit:
    st.subheader("依部門與單位查看各項預算")
    all_departments = sorted([x for x in current_expense_accounts["department"].dropna().unique() if x])
    selected_departments = st.multiselect("選擇部門", all_departments, default=all_departments)
    view_mode = st.radio("顯示方式", ["樹狀展開", "表格明細", "階層列總表"], horizontal=True)

    filtered = current_hierarchy_df.copy()
    if selected_departments:
        filtered = filtered[filtered["department"].isin(selected_departments) | filtered["level"].eq("合計")]

    if view_mode == "樹狀展開":
        render_budget_tree(filtered)
    elif view_mode == "表格明細":
        subject_df = filtered[filtered["level"] == "科目"].copy()
        display_df(
            subject_df,
            {
                "department": "部門", "unit": "單位", "program_group": "計畫群", "program": "計畫",
                "account_code": "代碼", "account_name": "科目", "amount": "金額", "description": "說明",
                "review_result_suggestion": "審查結果之建議", "path": "完整路徑",
            },
            amount_cols=["amount"],
        )
    else:
        display_df(
            filtered,
            {
                "level": "層級", "款": "款", "項": "項", "目": "目", "節": "節", "department": "部門",
                "unit": "單位", "program_group": "計畫群", "program": "計畫", "account_code": "代碼",
                "account_name": "名稱", "amount": "金額", "review_result_suggestion": "審查結果之建議", "path": "完整路徑",
            },
            amount_cols=["amount"],
        )


# =========================================================
# 科目層級瀏覽
# =========================================================
with tab_account:
    st.subheader("依科目查看各單位與計畫預算")
    browse_df = current_expense_accounts.copy()
    browse_df["account_label"] = browse_df["account_code"].fillna("") + "｜" + browse_df["account_name"].fillna("")
    options = sorted([x for x in browse_df["account_label"].dropna().unique() if x.strip() != "｜"])
    selected_accounts = st.multiselect("選擇科目", options, default=options[: min(5, len(options))])
    mode = st.radio("科目瀏覽方式", ["科目彙總", "科目 → 單位 → 計畫明細"], horizontal=True)
    if selected_accounts:
        browse_df = browse_df[browse_df["account_label"].isin(selected_accounts)]

    if mode == "科目彙總":
        account_summary = browse_df.groupby(["account_code", "account_name"], dropna=False)["amount"].sum().reset_index().sort_values("amount", ascending=False)
        display_df(account_summary, {"account_code": "代碼", "account_name": "科目", "amount": "本期預算合計"}, amount_cols=["amount"])
        if not account_summary.empty:
            chart_df = account_summary.copy()
            chart_df["label"] = chart_df["account_code"].fillna("") + "｜" + chart_df["account_name"].fillna("")
            st.bar_chart(chart_df.set_index("label")["amount"])
    else:
        for account_label in selected_accounts:
            one = browse_df[browse_df["account_label"] == account_label].copy()
            if one.empty:
                continue
            with st.expander(f"{account_label}｜合計 {format_amount(one['amount'].sum())} 元", expanded=False):
                unit_summary = one.groupby(["department", "unit"], dropna=False)["amount"].sum().reset_index().sort_values("amount", ascending=False)
                st.markdown("#### 各單位合計")
                display_df(unit_summary, {"department": "部門", "unit": "單位", "amount": "該科目預算合計"}, amount_cols=["amount"])
                st.markdown("#### 計畫明細")
                display_df(
                    one,
                    {
                        "department": "部門", "unit": "單位", "program_group": "計畫群", "program": "計畫",
                        "amount": "金額", "description": "說明", "review_result_suggestion": "審查結果之建議", "path": "完整路徑",
                    },
                    amount_cols=["amount"],
                )


# =========================================================
# 跨年度比較
# =========================================================
with tab_year_compare:
    st.subheader("跨年度比較：各年度各單位各科目預算變化")

    if multi_year_budget_df.empty:
        st.warning("尚未成功讀取多年度預算資料。請在左側上傳多年度預算 Excel 檔。")
    else:
        compare_df = multi_year_budget_df.copy()

        all_periods = sorted(compare_df["period_label"].dropna().unique())
        all_review_units = sorted(compare_df["review_unit"].dropna().unique())
        all_accounts = sorted(compare_df["account_label"].dropna().unique())

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            selected_periods = st.multiselect("選擇年度／期別", all_periods, default=all_periods)
        with col_b:
            selected_review_units = st.multiselect(
                "選擇審查單位／小組",
                all_review_units,
                default=all_review_units[: min(8, len(all_review_units))],
            )
        with col_c:
            default_accounts = [x for x in all_accounts if "工作費" in x]
            selected_accounts = st.multiselect(
                "選擇科目",
                all_accounts,
                default=default_accounts[:1] if default_accounts else all_accounts[: min(3, len(all_accounts))],
            )

        filtered_year_df = compare_df.copy()
        if selected_periods:
            filtered_year_df = filtered_year_df[filtered_year_df["period_label"].isin(selected_periods)]
        if selected_review_units:
            filtered_year_df = filtered_year_df[filtered_year_df["review_unit"].isin(selected_review_units)]
        if selected_accounts:
            filtered_year_df = filtered_year_df[filtered_year_df["account_label"].isin(selected_accounts)]

        summary = (
            filtered_year_df.groupby(
                [
                    "period_label",
                    "department",
                    "review_unit",
                    "review_unit_type",
                    "account_code",
                    "account_name",
                    "account_label",
                ],
                dropna=False,
            )["amount"]
            .sum()
            .reset_index()
            .rename(columns={"amount": "budget_amount"})
            .sort_values(["account_label", "review_unit", "period_label"])
        )

        st.markdown("### 明細表：年度 × 審查單位／小組 × 科目")
        display_df(
            summary,
            {
                "period_label": "年度／期別",
                "department": "部門",
                "review_unit": "審查單位／小組",
                "review_unit_type": "類型",
                "account_code": "代碼",
                "account_name": "科目",
                "budget_amount": "預算數",
            },
            amount_cols=["budget_amount"],
        )

        st.markdown("### 樞紐表：各年度預算變化")
        if summary.empty:
            st.info("目前篩選條件下沒有資料。")
        else:
            pivot = summary.pivot_table(
                index=["department", "review_unit", "review_unit_type", "account_code", "account_name"],
                columns="period_label",
                values="budget_amount",
                aggfunc="sum",
                fill_value=0,
            ).reset_index()

            period_cols = [
                col
                for col in pivot.columns
                if col not in ["department", "review_unit", "review_unit_type", "account_code", "account_name"]
            ]

            if len(period_cols) >= 2:
                first_period = period_cols[0]
                last_period = period_cols[-1]
                pivot["差額"] = pivot[last_period] - pivot[first_period]
                pivot["成長率"] = pivot.apply(
                    lambda row: None if row[first_period] == 0 else row["差額"] / row[first_period],
                    axis=1,
                )

            display_pivot = pivot.copy()
            for col in period_cols:
                display_pivot[col] = display_pivot[col].apply(format_amount)
            if "差額" in display_pivot.columns:
                display_pivot["差額"] = display_pivot["差額"].apply(format_amount)
            if "成長率" in display_pivot.columns:
                display_pivot["成長率"] = display_pivot["成長率"].apply(
                    lambda x: "上期為 0，無法計算" if x is None or pd.isna(x) else format_rate(x)
                )

            st.dataframe(
                display_pivot.rename(
                    columns={
                        "department": "部門",
                        "review_unit": "審查單位／小組",
                        "review_unit_type": "類型",
                        "account_code": "代碼",
                        "account_name": "科目",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

            st.markdown("### 趨勢圖")
            chart_summary = (
                summary.groupby(["period_label", "review_unit", "account_name"], dropna=False)["budget_amount"]
                .sum()
                .reset_index()
            )
            chart_summary["label"] = chart_summary["review_unit"] + "｜" + chart_summary["account_name"]

            if not chart_summary.empty:
                chart_data = chart_summary.pivot_table(
                    index="period_label",
                    columns="label",
                    values="budget_amount",
                    aggfunc="sum",
                    fill_value=0,
                )
                st.line_chart(chart_data)

            csv_data = summary.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "下載跨年度比較明細 CSV",
                data=csv_data,
                file_name="multi_year_budget_comparison.csv",
                mime="text/csv",
            )


# =========================================================
# 審查規則
# =========================================================
with tab_rules:
    st.subheader("核心審查規則")

    if previous_final_file is None or previous_final_accounts.empty:
        st.warning("請在左側上傳上期決算 Excel 檔，才能執行審查規則。")
    else:
        rule1 = rule_1_personnel_cap(
            current_expense_accounts_raw,
            previous_final_accounts,
            current_income_df,
        )
        rule3 = rule_3_low_execution_freeze(
            current_expense_accounts_raw,
            previous_final_accounts,
        )

        m1, m2, m3 = st.columns(3)
        m1.metric("人事合理性審查觸發", int(rule1["violate_rule"].sum()))
        m2.metric("低執行率人事科目觸發", len(rule3))
        freeze_total = rule1["suggested_freeze_amount"].sum() + (
            rule3["suggested_freeze_amount"].sum() if not rule3.empty else 0
        )
        m3.metric("建議凍結總額", f"{freeze_total:,.0f} 元")

        st.markdown("### 規則一：人事相關費用合理性審查")
        st.markdown(
            """
            本規則檢查工作費、臨時工資、稿費、製圖費四項人事相關費用。

            ```text
            上限 A = 上期人事執行決算 × 120%
            上限 B = 上期人事執行決算 × (1 + 整體預算成長率 × 110%)
            合理人事費上限 = max(上限 A, 上限 B) + 本期新增非會費收入
            建議凍結金額 = max(0, 本期人事相關預算 - 合理人事費上限)
            ```
            """
        )

        display_df(
            rule1,
            {
                "department": "部門",
                "review_unit": "審查單位／小組",
                "review_unit_type": "類型",
                "current_personnel_budget": "本期人事相關預算",
                "previous_personnel_final": "上期人事執行決算",
                "current_total_budget": "本期總預算",
                "previous_total_final": "上期總決算",
                "total_growth_rate": "整體預算成長率",
                "cap_a_previous_120": "上限 A：上期人事決算 120%",
                "cap_b_growth_adjusted": "上限 B：成長調整上限",
                "non_fee_income": "本期新增非會費收入",
                "reasonable_personnel_cap": "合理人事費上限",
                "suggested_freeze_amount": "建議凍結金額",
                "violate_rule": "是否觸發",
                "risk_reason": "審查理由",
            },
            amount_cols=[
                "current_personnel_budget",
                "previous_personnel_final",
                "current_total_budget",
                "previous_total_final",
                "cap_a_previous_120",
                "cap_b_growth_adjusted",
                "non_fee_income",
                "reasonable_personnel_cap",
                "suggested_freeze_amount",
            ],
            rate_cols=["total_growth_rate"],
        )

        st.markdown("### 規則二：低執行率人事科目凍結建議")
        st.markdown(
            """
            本規則只檢查工作費、臨時工資、稿費、製圖費四項人事相關費用，並以「審查單位／小組 × 科目」作為比較粒度。

            ```text
            若 上期同單位同人事科目決算合計 / 上期同單位同人事科目預算合計 < 80%
            且 本期同單位同人事科目預算合計 > 上期同單位同人事科目決算合計
            則 建議凍結金額 = 本期同單位同人事科目預算合計 - 上期同單位同人事科目決算合計
            ```
            """
        )

        if rule3.empty:
            st.success("目前沒有觸發規則二的項目。")
        else:
            rule3_display = rule3.copy()
            rule3_display["previous_execution_rate_display"] = rule3_display[
                "previous_execution_rate"
            ].apply(format_rate)

            display_df(
                rule3_display,
                {
                    "department": "部門",
                    "review_unit": "審查單位／小組",
                    "review_unit_type": "類型",
                    "current_programs": "本期涉及計畫",
                    "account_code": "代碼",
                    "account_name": "科目",
                    "current_budget_amount": "本期預算合計",
                    "previous_budget_amount": "上期預算合計",
                    "previous_final_amount": "上期決算合計",
                    "previous_execution_rate_display": "上期執行率",
                    "suggested_freeze_amount": "建議凍結金額",
                    "risk_reason": "審查理由",
                    "current_paths": "本期完整路徑",
                },
                amount_cols=[
                    "current_budget_amount",
                    "previous_budget_amount",
                    "previous_final_amount",
                    "suggested_freeze_amount",
                ],
            )

            csv_data = rule3_display.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "下載低執行率人事科目凍結建議 CSV",
                data=csv_data,
                file_name="low_execution_personnel_freeze_suggestions.csv",
                mime="text/csv",
            )


# =========================================================
# 搜尋
# =========================================================
with tab_search:
    st.subheader("關鍵字搜尋")
    keyword = st.text_input("輸入關鍵字，例如：工作費、行政費、講座、永續組")
    selected_units = st.multiselect("篩選單位", sorted([x for x in current_expense_accounts["unit"].dropna().unique() if x]))
    min_amount = st.number_input("最低金額", min_value=0, value=0, step=1000)

    filtered = current_expense_accounts.copy()
    if keyword.strip():
        filtered = linear_search(filtered, keyword)
    if selected_units:
        filtered = filtered[filtered["unit"].isin(selected_units)]
    filtered = filtered[filtered["amount"] >= min_amount].sort_values("amount", ascending=False)

    st.write(f"共找到 **{len(filtered)}** 筆結果。")
    display_df(
        filtered,
        {
            "department": "部門", "unit": "單位", "program_group": "計畫群", "program": "計畫",
            "account_code": "代碼", "account_name": "科目", "amount": "金額", "description": "說明",
            "review_result_suggestion": "審查結果之建議", "path": "完整路徑",
        },
        amount_cols=["amount"],
    )


# =========================================================
# 高金額排序
# =========================================================
with tab_ranking:
    st.subheader("高金額預算科目排序")
    top_k = st.slider("顯示前 K 筆", min_value=5, max_value=50, value=10)
    top_df = current_expense_accounts.sort_values("amount", ascending=False).head(top_k)
    display_df(
        top_df,
        {
            "department": "部門", "unit": "單位", "program": "計畫", "account_code": "代碼",
            "account_name": "科目", "amount": "金額", "description": "說明", "review_result_suggestion": "審查結果之建議",
        },
        amount_cols=["amount"],
    )
    if not top_df.empty:
        chart_df = top_df.copy()
        chart_df["label"] = chart_df["unit"] + "｜" + chart_df["program"] + "｜" + chart_df["account_name"]
        st.bar_chart(chart_df.set_index("label")["amount"])


# =========================================================
# 效能比較
# =========================================================
with tab_performance:
    st.subheader("資料結構效能比較：Linear Search vs Hash Map Index")
    st.markdown("這個區塊用同一份預算資料比較線性搜尋與 Hash Map / Inverted Index 搜尋。")
    performance_keyword = st.text_input("效能測試關鍵字", value="工作費")
    repeat = st.slider("重複查詢次數", min_value=10, max_value=1000, value=100, step=10)

    if st.button("執行效能比較"):
        result = measure_search_performance(current_expense_accounts, performance_keyword, repeat)
        if result is None:
            st.warning("請輸入關鍵字後再執行。")
        else:
            c1, c2, c3 = st.columns(3)
            c1.metric("Linear Search 總時間", f"{result['linear_time']:.6f} 秒")
            c2.metric("建立 Hash Index 時間", f"{result['index_build_time']:.6f} 秒")
            c3.metric("Hash Search 總時間", f"{result['hash_time']:.6f} 秒")

            comparison = pd.DataFrame([
                {"方法": "Linear Search", "總時間（秒）": result["linear_time"], "平均每次查詢（秒）": result["linear_time"] / result["repeat"], "找到筆數": result["linear_result_count"]},
                {"方法": "Hash Search（不含建 index）", "總時間（秒）": result["hash_time"], "平均每次查詢（秒）": result["hash_time"] / result["repeat"], "找到筆數": result["hash_result_count"]},
                {"方法": "Hash Search（含建 index）", "總時間（秒）": result["index_build_time"] + result["hash_time"], "平均每次查詢（秒）": (result["index_build_time"] + result["hash_time"]) / result["repeat"], "找到筆數": result["hash_result_count"]},
            ])
            st.dataframe(comparison, use_container_width=True, hide_index=True)
            st.bar_chart(comparison.set_index("方法")["總時間（秒）"])


# =========================================================
# 清理後資料
# =========================================================
with tab_raw:
    st.subheader("清理後資料")
    source = st.radio("資料來源", ["本期支出預算", "本期收入預算", "上期支出決算"], horizontal=True)
    scope = st.radio("資料範圍", ["只顯示科目列", "顯示完整階層"], horizontal=True)

    if source == "本期支出預算":
        raw_full = current_hierarchy_df
        raw_accounts = current_expense_accounts
    elif source == "本期收入預算":
        raw_full = current_income_df
        raw_accounts = current_income_df[current_income_df["level"] == "科目"].copy()
    else:
        if previous_final_df.empty:
            st.warning("尚未上傳或成功解析上期決算表。")
            st.stop()
        raw_full = previous_final_df
        raw_accounts = previous_final_accounts

    raw_display = raw_full if scope == "顯示完整階層" else raw_accounts
    st.dataframe(raw_display.drop(columns=["search_text"], errors="ignore"), use_container_width=True, hide_index=True)

    csv_data = raw_display.drop(columns=["search_text"], errors="ignore").to_csv(index=False).encode("utf-8-sig")
    st.download_button("下載清理後 CSV", data=csv_data, file_name="cleaned_budget_data.csv", mime="text/csv")
