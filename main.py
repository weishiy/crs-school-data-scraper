import os
import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from pypinyin import lazy_pinyin
import re


BASE_URL = "https://www.crs.jsj.edu.cn"
LIST_URL = "https://www.crs.jsj.edu.cn/aproval/orglists"
REFERER = "https://www.crs.jsj.edu.cn/index/sort/1006"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Referer": REFERER,
    "Origin": BASE_URL,
    "Content-Type": "application/x-www-form-urlencoded",
}

# 先只跑一个国家测试
COUNTRIES = ["英国"]

# 985、211 相关的高校名单（用于辅助提取大学名称）
# 这里直接写死了，后续可以改成从文件加载
University_985 = {
    "北京大学", "中国人民大学", "清华大学", "北京航空航天大学", "北京理工大学",
    "中国农业大学", "北京师范大学", "中央民族大学",
    "南开大学", "天津大学",
    "大连理工大学", "东北大学",
    "吉林大学",
    "哈尔滨工业大学",
    "复旦大学", "同济大学", "上海交通大学", "华东师范大学",
    "南京大学", "东南大学",
    "浙江大学",
    "中国科学技术大学",
    "厦门大学",
    "山东大学", "中国海洋大学",
    "武汉大学", "华中科技大学",
    "中南大学", "湖南大学", "国防科技大学",
    "中山大学", "华南理工大学",
    "四川大学", "电子科技大学",
    "重庆大学",
    "西安交通大学", "西北工业大学",
    "西北农林科技大学",
    "兰州大学",
}

University_211 = {
    "北京大学", "中国人民大学", "清华大学", "北京交通大学", "北京工业大学", "北京航空航天大学",
    "北京理工大学", "北京科技大学", "北京化工大学", "北京邮电大学", "中国农业大学",
    "北京林业大学", "北京中医药大学", "北京师范大学", "北京外国语大学", "中国传媒大学",
    "中央财经大学", "对外经济贸易大学", "中国政法大学", "中央民族大学", "华北电力大学",
    "中国矿业大学", "中国石油大学", "中国地质大学",
    "南开大学", "天津大学", "天津医科大学",
    "河北工业大学",
    "太原理工大学",
    "内蒙古大学",
    "辽宁大学", "大连理工大学", "东北大学", "大连海事大学",
    "吉林大学", "延边大学", "东北师范大学",
    "哈尔滨工业大学", "哈尔滨工程大学", "东北农业大学", "东北林业大学",
    "复旦大学", "同济大学", "上海交通大学", "华东理工大学", "东华大学",
    "华东师范大学", "上海外国语大学", "上海财经大学",
    "苏州大学", "南京大学", "东南大学", "南京航空航天大学", "南京理工大学",
    "中国矿业大学", "河海大学", "江南大学", "南京农业大学", "中国药科大学", "南京师范大学",
    "浙江大学",
    "安徽大学", "中国科学技术大学",
    "福州大学", "厦门大学",
    "南昌大学",
    "山东大学", "中国海洋大学", "中国石油大学",
    "郑州大学",
    "武汉大学", "华中科技大学", "中国地质大学", "武汉理工大学", "华中农业大学",
    "华中师范大学", "中南财经政法大学",
    "湖南大学", "中南大学", "湖南师范大学", "国防科技大学",
    "中山大学", "暨南大学", "华南理工大学", "华南师范大学",
    "广西大学",
    "海南大学",
    "四川大学", "西南交通大学", "电子科技大学", "四川农业大学", "西南财经大学",
    "重庆大学", "西南大学",
    "贵州大学",
    "云南大学",
    "西藏大学",
    "西安交通大学", "西北工业大学", "西安电子科技大学", "长安大学",
    "西北大学", "陕西师范大学", "西北农林科技大学",
    "兰州大学",
    "青海大学",
    "宁夏大学",
    "新疆大学", "石河子大学",
}


# ========= 工具函数 =========
def to_pinyin(text: str) -> str:
    return "".join(lazy_pinyin(str(text)))


def extract_university_name(full_name: str, category: str) -> str:
    full_name = full_name.replace("●", "").strip()

    if category == "合作办学项目" and "与" in full_name:
        return full_name.split("与")[0].strip()

    school_suffixes = ["职业技术学院", "职业学院", "师范大学", "大学", "学院", "学校"]
    for suffix in school_suffixes:
        idx = full_name.find(suffix)
        if idx != -1:
            return full_name[: idx + len(suffix)].strip()

    return ""

def mark_985_211(university_name: str) -> dict:
    name = university_name.strip()

    is_985 = name in University_985
    is_211 = name in University_211

    return {
        "is_985": is_985,
        "is_211": is_211,
        "is_985_211": is_985 or is_211
    }


# ========= 一级爬虫：列表页 =========
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

        # 每个 li 是一条记录
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
            university_tags = mark_985_211(university_name)

            records.append(
                {
                    "country": country,
                    "region": region,
                    "category": category,
                    "university_name": university_name,
                    "name": full_text,
                    "link": link,
                    **university_tags,
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


# ========= 二级爬虫：详情页 =========
def fetch_detail_html(session: requests.Session, detail_url: str) -> str:
    response = session.get(detail_url, headers=HEADERS, timeout=30)
    response.raise_for_status()
    return response.text


def extract_detail_fields(html: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")

    detail_data = {
        "level": "",
        "duration": "",
        "major_or_course": "",
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

def search_contact_page_better(row):
    from ddgs import DDGS

    university = row.get("university_name", "").strip()
    project_name = row.get("name", "").strip()
    level = row.get("level", "").strip()

    short_project = project_name.replace("合作举办", " ").replace("教育项目", " ").strip()

    queries = [
        f'{project_name} 招生简章 site:edu.cn',
        f'{university} {short_project} 招生简章 site:edu.cn',
        f'{university} {short_project} 联系方式 site:edu.cn',
        f'{university} {short_project} {level} 招生 site:edu.cn',
        f'{university} {short_project} site:edu.cn',   # 兜底
    ]

    blocked_domains = [
        "zhihu.com", "sohu.com", "163.com", "baidu.com",
        "wenku.baidu.com", "douban.com", "bilibili.com",
        "xiaohongshu.com", "sina.com.cn", "qq.com"
    ]

    preferred_url_parts = [
        "yz.", "graduate", "admission", "zs", "yjs", "mba", "som", "sem"
    ]

    best_url = ""
    best_score = -999

    fallback_url = ""

    with DDGS() as ddgs:
        for query in queries:
            try:
                results = ddgs.text(query, max_results=8)
            except Exception as e:
                print(f"搜索失败: {query} -> {e}")
                continue

            for r in results:
                title = (r.get("title") or "").strip()
                body = (r.get("body") or "").strip()
                url = (r.get("href") or r.get("url") or "").strip()

                if not url:
                    continue
                if any(bad in url for bad in blocked_domains):
                    continue
                if ".edu.cn" not in url and ".ac.cn" not in url:
                    continue

                # 先记一个宽松兜底结果
                if not fallback_url:
                    fallback_url = url

                text = f"{title} {body}"
                score = 0

                if "招生简章" in text:
                    score += 10
                if "联系方式" in text or "联系我们" in text:
                    score += 6
                if "招生" in text or "报名" in text:
                    score += 3

                if university and university in text:
                    score += 4

                if project_name[:15] and project_name[:15] in text:
                    score += 3

                if level and level in text:
                    score += 1

                if any(x in url.lower() for x in preferred_url_parts):
                    score += 4

                # 不要扣分太狠，只轻微扣
                if any(x in url.lower() for x in ["list", "index"]):
                    score -= 1

                if score > best_score:
                    best_score = score
                    best_url = url

    # 先返回严格筛选结果；如果没有，就返回宽松兜底结果
    return best_url if best_url else fallback_url

def search_contact_page_debug(row):
    from ddgs import DDGS

    query = f'{row["university_name"]} {row["name"]} 招生简章 联系方式'
    print("搜索词:", query)

    with DDGS() as ddgs:
        results = ddgs.text(query, max_results=5)
        for i, r in enumerate(results, 1):
            print(f"{i}. 标题: {r.get('title')}")
            print(f"   链接: {r.get('href') or r.get('url')}")
            print(f"   摘要: {r.get('body')}")
            print("-" * 60)

def extract_contact_fields_from_html(html: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n", strip=True)

    result = {
        "contact_person": "",
        "phone": "",
        "email": "",
        "wechat": "",
        "wechat_official_account": "",
        "address": "",
    }

    patterns = {
        "contact_person": [
            r"联系人\s*[：:]\s*([^\n]+)",
            r"联\s*系\s*人\s*[：:]\s*([^\n]+)",
        ],
        "phone": [
            r"联系方式\s*[：:]\s*([^\n]*?(?:\d{3,4}-\d{7,8}|\d{11})[^\n]*)",
            r"联系电话\s*[：:]\s*([^\n]+)",
            r"咨询电话\s*[：:]\s*([^\n]+)",
            r"电话\s*[：:]\s*([^\n]+)",
        ],
        "email": [
            r"电子邮箱\s*[：:]\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})",
            r"邮箱\s*[：:]\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})",
            r"E-?mail\s*[：:]\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})",
            r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})",
        ],
        "wechat": [
            r"微信号\s*[：:]\s*([^\n]+)",
            r"微信\s*[：:]\s*([^\n]+)",
        ],
        "wechat_official_account": [
            r"微信公众号\s*[：:]\s*([^\n]+)",
            r"公众号\s*[：:]\s*([^\n]+)",
        ],
        "address": [
            r"联系地址\s*[：:]\s*([^\n]+)",
            r"地址\s*[：:]\s*([^\n]+)",
        ],
    }

    for field, field_patterns in patterns.items():
        for pattern in field_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                result[field] = match.group(1).strip()
                break

    return result

def auto_enrich_contacts_better(detail_records, session):
    enriched = []

    for i, row in enumerate(detail_records, 1):
        print(f"自动搜索 {i}/{len(detail_records)}: {row['name']}")

        new_row = row.copy()

        contact_url = search_contact_page_better(row)
        new_row["contact_url"] = contact_url

        # 默认空值
        new_row["contact_person"] = ""
        new_row["phone"] = ""
        new_row["email"] = ""
        new_row["wechat"] = ""
        new_row["wechat_official_account"] = ""
        new_row["address"] = ""

        if contact_url:
            try:
                html = fetch_detail_html(session, contact_url)
                contact_data = extract_contact_fields_from_html(html)
                print("URL:", contact_url)
                print("提取结果:", contact_data)
                new_row.update(contact_data)
            except Exception as e:
                print(f"联系方式解析失败: {contact_url} -> {e}")

        enriched.append(new_row)

    return enriched

# ========= Excel 格式 =========
def format_sheet(ws, df: pd.DataFrame):
    ws.freeze_panes = "A2"

    # 对齐 + 自动换行
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="center",
                horizontal="center",
                wrap_text=True
            )

    # 根据列名设置列宽
    col_index = {col_name: idx + 1 for idx, col_name in enumerate(df.columns)}

    def set_width(col_name: str, width: int):
        if col_name in col_index:
            col_letter = ws.cell(row=1, column=col_index[col_name]).column_letter
            ws.column_dimensions[col_letter].width = width

    set_width("country", 10)
    set_width("region", 10)
    set_width("category", 16)
    set_width("university_name", 22)
    set_width("name", 46)
    set_width("link", 16)
    set_width("level", 10)
    set_width("duration", 10)
    set_width("major_or_course", 20)
    set_width("contact_url", 40)
    set_width("contact_person", 18)
    set_width("phone", 18)
    set_width("email", 26)
    set_width("wechat", 22)
    set_width("wechat_official_account", 24)
    set_width("address", 32)

    # name 左对齐
    left_align_cols = [
    "name",
    "contact_url",
    "contact_person",
    "phone",
    "email",
    "wechat",
    "wechat_official_account",
    "address",
]

    for col_name in left_align_cols:
        if col_name in col_index:
            col_letter = ws.cell(row=1, column=col_index[col_name]).column_letter
            for row_num in range(2, ws.max_row + 1):
                ws[f"{col_letter}{row_num}"].alignment = Alignment(
                    vertical="center",
                    horizontal="left",
                    wrap_text=True
                )

    # link 变可点击
    if "link" in col_index:
        col_letter = ws.cell(row=1, column=col_index["link"]).column_letter
        for row_num in range(2, ws.max_row + 1):
            cell = ws[f"{col_letter}{row_num}"]
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
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 28

    # 按列名找位置
    country_col = col_index.get("country")
    region_col = col_index.get("region")
    category_col = col_index.get("category")

    # 合并 country
    if country_col:
        start = 2
        for _, group in df.groupby("country", sort=False):
            end = start + len(group) - 1
            if len(group) > 1:
                ws.merge_cells(start_row=start, start_column=country_col, end_row=end, end_column=country_col)
            start = end + 1

    # 合并 region（同一 country 内）
    if country_col and region_col:
        start = 2
        for _, group in df.groupby(["country", "region"], sort=False):
            end = start + len(group) - 1
            if len(group) > 1:
                ws.merge_cells(start_row=start, start_column=region_col, end_row=end, end_column=region_col)
            start = end + 1

    # 合并 category（同一 country + region 内）
    if country_col and region_col and category_col:
        start = 2
        for _, group in df.groupby(["country", "region", "category"], sort=False):
            end = start + len(group) - 1
            if len(group) > 1:
                ws.merge_cells(start_row=start, start_column=category_col, end_row=end, end_column=category_col)
            start = end + 1


# ========= 一个 Excel，多 tab 导出 =========
def save_all_to_excel(all_records: list[dict], detail_records: list[dict], filename="output/crs_full_data.xlsx"):
    df_all = pd.DataFrame(all_records)

    if df_all.empty:
        print("没有一级数据可导出。")
        return

    # 一级排序
    df_all["country_pinyin"] = df_all["country"].apply(to_pinyin)
    df_all["region_pinyin"] = df_all["region"].apply(to_pinyin)
    df_all["category_pinyin"] = df_all["category"].apply(to_pinyin)
    df_all["university_pinyin"] = df_all["university_name"].apply(to_pinyin)
    df_all["name_pinyin"] = df_all["name"].apply(to_pinyin)

    df_all = df_all.sort_values(
        by=["country_pinyin", "region_pinyin", "category_pinyin", "university_pinyin", "name_pinyin"]
    ).reset_index(drop=True)

    df_all = df_all.drop(columns=["country_pinyin", "region_pinyin", "category_pinyin", "university_pinyin", "name_pinyin"])
    df_all = df_all[
        [
            "country", 
            "region", 
            "category", 
            "university_name", 
            "is_985", 
            "is_211",
            "is_985_211",
            "name", 
            "link"]]

    df_projects = df_all[df_all["category"] == "合作办学项目"].copy().reset_index(drop=True)

    df_non_985_211_projects = df_all[
    (df_all["category"] == "合作办学项目") &
    (df_all["is_985_211"] == False)
    ].copy().reset_index(drop=True)

    # 二级详情数据
    df_detail = pd.DataFrame(detail_records)
    if not df_detail.empty:
        # 补排序
        df_detail["country_pinyin"] = df_detail["country"].apply(to_pinyin)
        df_detail["region_pinyin"] = df_detail["region"].apply(to_pinyin)
        df_detail["category_pinyin"] = df_detail["category"].apply(to_pinyin)
        df_detail["university_pinyin"] = df_detail["university_name"].apply(to_pinyin)
        df_detail["name_pinyin"] = df_detail["name"].apply(to_pinyin)

        df_detail = df_detail.sort_values(
            by=["country_pinyin", "region_pinyin", "category_pinyin", "university_pinyin", "name_pinyin"]
        ).reset_index(drop=True)

        df_detail = df_detail.drop(columns=["country_pinyin", "region_pinyin", "category_pinyin", "university_pinyin", "name_pinyin"])

        detail_cols = [
            "country",
            "region",
            "category",
            "university_name",
            "is_985",
            "is_211",
            "is_985_211",
            "name",
            "level",
            "duration",
            "major_or_course",
            "link",
        ]
        existing_cols = [c for c in detail_cols if c in df_detail.columns]
        df_detail = df_detail[existing_cols]

    try:
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df_all.to_excel(writer, sheet_name="全部数据", index=False)
            df_projects.to_excel(writer, sheet_name="合作办学项目", index=False)
            df_non_985_211_projects.to_excel(writer, sheet_name="非985_211合作办学项目", index=False)
            if not df_detail.empty:
                df_detail.to_excel(writer, sheet_name="项目详情", index=False)
    except PermissionError:
        filename = "output/crs_full_data_new.xlsx"
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df_all.to_excel(writer, sheet_name="全部数据", index=False)
            df_projects.to_excel(writer, sheet_name="合作办学项目", index=False)
            df_non_985_211_projects.to_excel(writer, sheet_name="非985_211合作办学项目", index=False)
            if not df_detail.empty:
                df_detail.to_excel(writer, sheet_name="项目详情", index=False)
        print(f"原文件被占用，已改存为: {filename}")


    wb = load_workbook(filename)

    ws_all = wb["全部数据"]
    format_sheet(ws_all, df_all)

    ws_projects = wb["合作办学项目"]
    format_sheet(ws_projects, df_projects)

    ws_non = wb["非985_211合作办学项目"]
    format_sheet(ws_non, df_non_985_211_projects)
    
    if not df_detail.empty:
        ws_detail = wb["项目详情"]
        format_sheet(ws_detail, df_detail)

    wb.save(filename)
    print(f"Excel 已导出: {filename}")


# ========= 主流程 =========
def main():
    os.makedirs("output", exist_ok=True)

    session = requests.Session()
    session.get(REFERER, headers=HEADERS, timeout=30)

    all_records = []

    # 1. 一级爬虫
    for country in COUNTRIES:
        print(f"正在抓取列表页: {country}")
        html = fetch_country_html(session, country)
        records = extract_records(html, country)
        print(f"{country} 抓到 {len(records)} 条一级数据")
        all_records.extend(records)

    if not all_records:
        print("没有抓到一级数据。")
        return

    # 2. 从一级数据里筛出合作办学项目
    project_records = [
        r for r in all_records 
        if r["category"] == "合作办学项目" and not r["is_985_211"]
        ]
    
    print(f"合作办学项目共有 {len(project_records)} 条")

    # 3. 二级爬虫：只抓合作办学项目详情
    detail_records = []

    # 先测试前10条；没问题再改成 project_records
    for i, row in enumerate(project_records):
        print(f"正在抓详情页 {i}: {row['link']}")
        try:
            detail_html = fetch_detail_html(session, row["link"])
            detail = extract_detail_fields(detail_html)
            merged = {**row, **detail}
            detail_records.append(merged)
        except Exception as e:
            print(f"抓取失败: {row['link']} -> {e}")

    print(f"成功抓到 {len(detail_records)} 条二级详情数据")

    # 4. 导出一个 Excel，多 tab
    save_all_to_excel(all_records, detail_records)

    # ===== 自动补联系方式 =====
    enriched = auto_enrich_contacts_better(detail_records, session)  # 先测试5条

    # 写入新sheet
    df_contacts = pd.DataFrame(enriched)

    contact_cols = [
        "country",
        "region",
        "category",
        "university_name",
        "name",
        "level",
        "duration",
        "major_or_course",
        "contact_url",
        "contact_person",
        "phone",
        "email",
        "wechat",
        "wechat_official_account",
        "address",
        "link",
    ]

    existing_cols = [c for c in contact_cols if c in df_contacts.columns]
    df_contacts = df_contacts[existing_cols]

    with pd.ExcelWriter("output/crs_full_data.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_contacts.to_excel(writer, sheet_name="自动联系方式", index=False)

    wb = load_workbook("output/crs_full_data.xlsx")
    ws_contacts = wb["自动联系方式"]
    format_sheet(ws_contacts, df_contacts)
    wb.save("output/crs_full_data.xlsx")

    print("联系方式已写入")



if __name__ == "__main__":
    main()