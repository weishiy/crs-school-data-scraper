import os
import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from pypinyin import lazy_pinyin

BASE_URL = "https://www.crs.jsj.edu.cn"
LIST_URL = "https://www.crs.jsj.edu.cn/aproval/orglists"
REFERER = "https://www.crs.jsj.edu.cn/index/sort/1006"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Referer": REFERER,
    "Origin": BASE_URL,
    "Content-Type": "application/x-www-form-urlencoded",
}

# 先只跑一个国家，确认没问题后再加别的
COUNTRIES = ["英国"]


def to_pinyin(text):
    return "".join(lazy_pinyin(str(text)))


def extract_university_name(full_name: str, category: str) -> str:
    full_name = full_name.replace("●", "").strip()

    if category == "合作办学项目":
        if "与" in full_name:
            return full_name.split("与")[0].strip()

    school_suffixes = ["职业技术学院", "职业学院", "师范大学", "大学", "学院", "学校"]

    for suffix in school_suffixes:
        idx = full_name.find(suffix)
        if idx != -1:
            return full_name[: idx + len(suffix)].strip()

    return ""


def fetch_country_html(session: requests.Session, country: str) -> str:
    data = {
        "subjdirect": "",
        "cn_runschool": "",
        "en_runschool": "",
        "local": country,
    }
    response = session.post(LIST_URL, headers=HEADERS, data=data, timeout=30)
    response.raise_for_status()
    return response.text


def extract_records(html: str, country: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    records = []
    current_region = None

    for tr in soup.find_all("tr"):
        tds = tr.find_all("td")
        if not tds:
            continue

        texts = [td.get_text(" ", strip=True) for td in tds]

        # 跳过表头
        if "地区" in texts and "项目/机构" in texts:
            continue

        # 有地区列
        if len(tds) >= 3:
            region = texts[0]
            category = texts[1]
            name_td = tds[2]

            if not region or not category:
                continue

            current_region = region

        # rowspan 续行
        elif len(tds) == 2 and current_region:
            region = current_region
            category = texts[0]
            name_td = tds[1]

            if not category:
                continue
        else:
            continue

        # 每个 li 是一条完整记录
        lis = name_td.find_all("li")
        for li in lis:
            full_text = li.get_text(" ", strip=True).replace("●", "").strip()
            if not full_text:
                continue

            link = ""
            for a in li.find_all("a", href=True):
                href = a.get("href")
                if href and "/aproval/detail/" in href:
                    link = urljoin(BASE_URL, href).strip()
                    break

            if not link:
                continue

            university_name = extract_university_name(full_text, category)

            records.append(
                {
                    "country": country,
                    "region": region,
                    "category": category,
                    "university_name": university_name,
                    "name": full_text,
                    "link": link,
                }
            )

    # 去重
    unique_records = []
    seen = set()
    for row in records:
        key = (
            row["country"],
            row["region"],
            row["category"],
            row["university_name"],
            row["name"],
            row["link"],
        )
        if key not in seen:
            seen.add(key)
            unique_records.append(row)

    return unique_records


def format_sheet(ws, df_sheet):
    ws.freeze_panes = "A2"

    # 列宽
    ws.column_dimensions["A"].width = 10   # region
    ws.column_dimensions["B"].width = 16   # category
    ws.column_dimensions["C"].width = 22   # university_name
    ws.column_dimensions["D"].width = 46   # name
    ws.column_dimensions["E"].width = 16   # link

    # 全部先居中+换行
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="center",
                horizontal="center",
                wrap_text=True
            )

    # name 左对齐
    for row in range(2, ws.max_row + 1):
        ws[f"D{row}"].alignment = Alignment(
            vertical="center",
            horizontal="left",
            wrap_text=True
        )

    # link 做成可点击
    for row in range(2, ws.max_row + 1):
        cell = ws[f"E{row}"]
        url = cell.value
        if url and isinstance(url, str):
            cell.value = "查看详情"
            cell.hyperlink = url
            cell.style = "Hyperlink"
            cell.alignment = Alignment(
                vertical="center",
                horizontal="center",
                wrap_text=True
            )

    # 行高
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 28

    # 最后 merge
    start_row = 2
    for _, group in df_sheet.groupby("region", sort=False):
        end_row = start_row + len(group) - 1
        if len(group) > 1:
            ws.merge_cells(
                start_row=start_row,
                start_column=1,
                end_row=end_row,
                end_column=1
            )
        start_row = end_row + 1

    start_row = 2
    for _, group in df_sheet.groupby(["region", "category"], sort=False):
        end_row = start_row + len(group) - 1
        if len(group) > 1:
            ws.merge_cells(
                start_row=start_row,
                start_column=2,
                end_row=end_row,
                end_column=2
            )
        start_row = end_row + 1


def save_to_excel(all_records: list[dict], filename: str = "output/crs_school_list.xlsx") -> None:
    df = pd.DataFrame(all_records)

    if df.empty:
        print("没有数据可导出。")
        return

    # 拼音排序
    df["region_pinyin"] = df["region"].apply(to_pinyin)
    df["category_pinyin"] = df["category"].apply(to_pinyin)
    df["university_pinyin"] = df["university_name"].apply(to_pinyin)
    df["name_pinyin"] = df["name"].apply(to_pinyin)

    df = df.sort_values(
        by=["region_pinyin", "category_pinyin", "university_pinyin", "name_pinyin"]
    ).reset_index(drop=True)

    df = df.drop(columns=["region_pinyin", "category_pinyin", "university_pinyin", "name_pinyin"])

    df = df[["region", "category", "university_name", "name", "link"]]

    df_projects = df[df["category"] == "合作办学项目"].copy().reset_index(drop=True)

    try:
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="全部数据", index=False)
            df_projects.to_excel(writer, sheet_name="合作办学项目", index=False)
    except PermissionError:
        filename = "output/crs_school_list_new.xlsx"
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="全部数据", index=False)
            df_projects.to_excel(writer, sheet_name="合作办学项目", index=False)
        print(f"原文件被占用，已改存为: {filename}")

    wb = load_workbook(filename)

    ws_all = wb["全部数据"]
    format_sheet(ws_all, df)

    ws_projects = wb["合作办学项目"]
    format_sheet(ws_projects, df_projects)

    wb.save(filename)
    print(f"列表页 Excel 已导出: {filename}")


def fetch_detail_html(session: requests.Session, detail_url: str) -> str:
    response = session.get(detail_url, headers=HEADERS, timeout=30)
    response.raise_for_status()
    return response.text


def extract_detail_fields(html: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")

    detail_data = {
        "level": "",
        "duration": "",
        "major_or_course": ""
    }

    for tr in soup.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        texts = [cell.get_text(" ", strip=True) for cell in cells if cell.get_text(" ", strip=True)]

        if len(texts) < 2:
            continue

        for i in range(0, len(texts) - 1, 2):
            key = texts[i].strip()
            value = texts[i + 1].strip()

            if key == "办学层次和类别":
                if "硕士" in value:
                    detail_data["level"] = "硕士"
                elif "本科" in value:
                    detail_data["level"] = "本科"
                elif "专科" in value:
                    detail_data["level"] = "专科"
                elif "博士" in value:
                    detail_data["level"] = "博士"
                else:
                    detail_data["level"] = value

            elif key == "学制":
                detail_data["duration"] = value

            elif key == "开设专业或课程":
                detail_data["major_or_course"] = value

    return detail_data


def save_detail_excel(detail_records: list[dict], filename: str = "output/crs_detail_data.xlsx") -> None:
    if not detail_records:
        print("没有详情数据可导出。")
        return

    df = pd.DataFrame(detail_records)

    preferred_cols = [
        "country",
        "region",
        "category",
        "university_name",
        "name",
        "level",
        "duration",
        "major_or_course",
        "link",
    ]

    existing_cols = [col for col in preferred_cols if col in df.columns]
    df = df[existing_cols]

    try:
        df.to_excel(filename, index=False)
    except PermissionError:
        filename = "output/crs_detail_data_new.xlsx"
        df.to_excel(filename, index=False)
        print(f"原详情文件被占用，已改存为: {filename}")

    print(f"详情页 Excel 已导出: {filename}")


def main():
    os.makedirs("output", exist_ok=True)

    session = requests.Session()
    session.get(REFERER, headers=HEADERS, timeout=30)

    all_records = []

    # 1. 抓列表页
    for country in COUNTRIES:
        print(f"正在抓取列表页: {country}")
        html = fetch_country_html(session, country)
        records = extract_records(html, country)
        print(f"{country} 抓到 {len(records)} 条列表数据")
        all_records.extend(records)

    if not all_records:
        print("没有抓到列表数据，请检查页面结构。")
        return

    # 2. 导出列表页 Excel
    save_to_excel(all_records)

    # 3. 抓详情页（先全量；如果怕慢，可改成 all_records[:10]）
    detail_records = []

    for i, row in enumerate(all_records[:10], 1):
        print(f"正在抓详情页 {i}/{len(all_records)}: {row['link']}")
        try:
            detail_html = fetch_detail_html(session, row["link"])
            detail_data = extract_detail_fields(detail_html)

            merged_row = {
                "country": row["country"],
                "region": row["region"],
                "category": row["category"],
                "university_name": row["university_name"],
                "name": row["name"],
                "link": row["link"],
                **detail_data
            }
            detail_records.append(merged_row)

        except Exception as e:
            print(f"抓取失败: {row['link']} -> {e}")

    # 4. 导出详情页 Excel
    save_detail_excel(detail_records)

    print("完成。")


if __name__ == "__main__":
    main()