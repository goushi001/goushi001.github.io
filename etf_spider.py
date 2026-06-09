import os
import time
import re
import requests
import pandas as pd
from io import BytesIO
from datetime import datetime, date, timedelta
from chinese_calendar import is_workday
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook


# ================= 配置区 =================
# GitHub Actions 环境：脚本与 Excel 在同一仓库
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_PATH = os.path.join(SCRIPT_DIR, "assets", "posts", "ETF份额统计.xlsx")

SSE_LIST = [
    "510300", "510310", "510050", "510500", "512100", "512400", "516650",
    "517520", "512480", "588200", "588000", "515070", "512660", "562500",
    "512170", "512010", "515790", "516970", "512200", "516510", "512880",
    "512800", "513090", "513180", "516150", "512690", "515220", "561360",
    "510170", "512890", "560080", "511130", "511260"
]

SZSE_LIST = [
    '159915', '159919', '159326', '159852', '159870',
    '159516', '159819', '159206', '159992', '159755',
    '159745', '159611', '159851', '159920', '159567', '159928', '159998', '159985'
]

SEARCH_ROW = 13
DATA_START_ROW = 18
DATE_COL = 3
# ==========================================


def shift_formula_rows(formula, delta):
    """将公式中所有行号偏移 delta，如 =D19 +1 → =D20"""
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula

    def replace_row(match):
        return f"{match.group(1)}{int(match.group(2)) + delta}"

    return re.sub(r'([A-Z]+)(\d+)', replace_row, formula)


# ─────────────────────────────────────────
# 第一步：准备 Excel（腾行 + 写公式）
# ─────────────────────────────────────────
def prepare_excel_new_row():
    """
    不使用 insert_rows，手动将第18行内容搬到第19行，
    再把公式原样写回第18行，数字列留空等待填入新数据。
    """
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"❌ 找不到 Excel: {EXCEL_FILE_PATH}")
        return None

    try:
        wb = load_workbook(EXCEL_FILE_PATH)
        ws = wb.active
        print(f"📋 工作表: {ws.title}")

        # 扫描第13行，建立 代码→列 映射
        code_to_col = {}
        for col in range(1, ws.max_column + 1):
            val = ws.cell(SEARCH_ROW, col).value
            if val is not None:
                try:
                    code = str(int(float(str(val))))
                    if len(code) == 6:
                        code_to_col[code] = col
                except:
                    pass
        print(f" 找到 {len(code_to_col)} 个ETF代码映射")

        # 读取第18行快照
        row18_snapshot = {
            col: ws.cell(DATA_START_ROW, col).value
            for col in range(1, ws.max_column + 1)
        }

        # 将第18行内容搬到第19行（公式行号+1，其余原样）
        for col, val in row18_snapshot.items():
            if val is None:
                continue
            if isinstance(val, str) and val.startswith('='):
                ws.cell(DATA_START_ROW + 1, col).value = shift_formula_rows(val, +1)
            else:
                ws.cell(DATA_START_ROW + 1, col).value = val

        # 清空第18行，把公式原样写回（行号不变=天然正确）
        for col in range(1, ws.max_column + 1):
            ws.cell(DATA_START_ROW, col).value = None
        for col, val in row18_snapshot.items():
            if isinstance(val, str) and val.startswith('='):
                ws.cell(DATA_START_ROW, col).value = val

        wb.save(EXCEL_FILE_PATH)
        print("✅ 第一步完成：新行已腾出，公式已写入第18行")
        return code_to_col

    except Exception as e:
        print(f"❌ Excel 准备失败（请先关闭文件）: {e}")
        return None


# ─────────────────────────────────────────
# 第二步：抓取数据
# ─────────────────────────────────────────
def fetch_szse_data(code_list):
    url = "https://www.szse.cn/api/report/ShowReport?SHOWTYPE=xlsx&CATALOGID=1105&TABKEY=tab1"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    results_map = {}
    try:
        print(f"\n--- 正在同步深交所数据 ({datetime.now().strftime('%H:%M:%S')}) ---")
        response = requests.get(url, headers=headers, timeout=15)
        df = pd.read_excel(BytesIO(response.content), header=None)
        df_str = df.astype(str)
        for code in code_list:
            mask = df_str.apply(lambda row: row.str.contains(str(code)).any(), axis=1)
            target_row = df[mask]
            if not target_row.empty:
                shares_val = 0.0
                for val in target_row.iloc[0].values:
                    try:
                        temp_val = float(str(val).replace(',', '').strip())
                        if temp_val > 1000000:
                            shares_val = temp_val
                    except:
                        pass
                if shares_val > 0:
                    results_map[str(code)] = shares_val / 10000
                    print(f" ✅ 深交所 {code}: {shares_val/10000:.2f} 万份")
                else:
                    print(f" ⚠️ 深交所未找到: {code}")
    except Exception as e:
        print(f"❌ 深交所同步失败: {e}")
    return results_map


def extract_sse_data_dom(driver, code, max_retries=2):
    """读 DOM，返回 (份额万份, 日期字符串)"""
    for attempt in range(1, max_retries + 1):
        try:
            wait = WebDriverWait(driver, 15)
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div.js_fundSize table.table tbody tr")
            ))
            rows = driver.find_elements(
                By.CSS_SELECTOR, "div.js_fundSize table.table tbody tr"
            )
            if rows:
                tds = rows[0].find_elements(By.TAG_NAME, "td")
                if len(tds) >= 4:
                    raw_date = tds[0].text.strip()
                    raw_share = tds[-1].text.replace(',', '').strip()
                    share_val = float(raw_share)
                    if share_val > 0:
                        parts = raw_date.split("-")
                        date_fmt = f"{parts[0]}//{int(parts[1])}/{int(parts[2])}"
                        return share_val, date_fmt
        except Exception as e:
            if attempt < max_retries:
                time.sleep(2)
            else:
                print(f" ⚠️ DOM读取失败 {code}: {e}")
    return None, None


def fetch_sse_data(code_list):
    results_map = {}
    failed_list = []
    detected_date = None

    options = Options()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--log-level=3')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )

    print(f"\n🚀 开始上交所处理 (共 {len(code_list)} 只)...")
    try:
        for idx, code in enumerate(code_list, start=1):
            try:
                driver.get(
                    f"https://www.sse.com.cn/assortment/fund/list/etfinfo/basic/"
                    f"index.shtml?FUNDID={code}"
                )
                driver.execute_script("window.scrollTo(0, 1800);")
                share, date_str = extract_sse_data_dom(driver, code)
                if share:
                    results_map[str(code)] = share
                    if detected_date is None and date_str:
                        detected_date = date_str
                    print(f" [{idx}/{len(code_list)}] ✅ 沪市 {code}: {share:.2f} 万份 日期: {date_str}")
                else:
                    failed_list.append(code)
                    print(f" [{idx}/{len(code_list)}] ⚠️ 沪市 {code} 失败")
            except Exception as e:
                failed_list.append(code)
                print(f" [{idx}/{len(code_list)}] ❌ 沪市 {code} 异常: {e}")
    finally:
        driver.quit()

    if failed_list:
        print(f"\n⚠️ 未获取到数据的沪市ETF: {failed_list}")
    return results_map, detected_date


# ─────────────────────────────────────────
# 第三步：将数据填入第18行
# ─────────────────────────────────────────
def fill_excel_data(code_to_col, data_dict, data_date):
    try:
        wb = load_workbook(EXCEL_FILE_PATH)
        ws = wb.active

        # 填日期
        if data_date:
            ws.cell(DATA_START_ROW, DATE_COL).value = data_date
            print(f"\n 📅 日期（来自上交所）: {data_date}")
        else:
            today = datetime.now()
            date_fmt = f"{today.year}//{today.month}/{today.day}"
            ws.cell(DATA_START_ROW, DATE_COL).value = date_fmt
            print(f"\n 📅 日期（系统备用）: {date_fmt}")

        # 填份额（直接覆盖，有旧值也没关系）
        filled_count = 0
        for code, share_val in data_dict.items():
            code_str = str(code)
            if code_str in code_to_col:
                share_col = code_to_col[code_str] + 2
                ws.cell(DATA_START_ROW, share_col).value = float(share_val)
                filled_count += 1
                print(f" 🎯 {code_str} → col{share_col}: {float(share_val):.2f} 万份")
            else:
                print(f" ⚠️ {code_str} 未在第{SEARCH_ROW}行找到对应列")

        wb.save(EXCEL_FILE_PATH)
        print(f"\n✨ 全部完成！成功填入 {filled_count}/{len(data_dict)} 只ETF份额")

    except Exception as e:
        print(f"❌ 数据填入失败（请先关闭文件）: {e}")
        import traceback
        traceback.print_exc()


# ─────────────────────────────────────────
# 主程序
# ─────────────────────────────────────────
def run_integration():
    start_time = datetime.now()
    print(f"{'='*55}")
    print(f" ETF 份额采集 {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*55}")

    # 第一步：准备 Excel（腾行 + 写公式）
    print("\n【第一步】准备 Excel 新行...")
    code_to_col = prepare_excel_new_row()
    if code_to_col is None:
        print("❌ Excel 准备失败，程序终止")
        return

    # 第二步：抓取数据
    print("\n【第二步】抓取数据...")
    final_results = fetch_szse_data(SZSE_LIST)
    sse_results, data_date = fetch_sse_data(SSE_LIST)
    final_results.update(sse_results)

    # 第三步：填入数据
    print("\n【第三步】填入数据...")
    fill_excel_data(code_to_col, final_results, data_date)

    elapsed = (datetime.now() - start_time).seconds
    print(f"\n📊 采集汇总: 成功 {len(final_results)}/{len(SSE_LIST)+len(SZSE_LIST)} 只")
    print(f"⏱️ 总耗时: {elapsed // 60} 分 {elapsed % 60} 秒")
    print(f"{'='*55}")


if __name__ == "__main__":
    # ── 前置判断：昨天是否为A股交易日 ──
    yesterday = date.today() - timedelta(days=1)
    if not is_workday(yesterday):
        print(f"⏭️ {yesterday} 非A股交易日（周末或法定节假日），今日不运行，跳过。")
        exit(0)
    else:
        print(f"✅ {yesterday} 是交易日，开始采集流程...")

    run_integration()
