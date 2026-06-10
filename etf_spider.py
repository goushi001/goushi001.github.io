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

SEARCH_ROW     = 13
DATA_START_ROW = 18
DATE_COL       = 3

ROW_DAILY     = 10
ROW_WEEKLY    = 9
ROW_MONTHLY   = 8
ROW_QUARTERLY = 7

# 历史数据区域终止行（该行会被覆盖，其余行完整保留）
MAX_HISTORY_ROW = 163
# ==========================================


def check_trading_day():
    yesterday = date.today() - timedelta(days=1)
    if not is_workday(yesterday):
        print(f"{'='*55}")
        print(f"  ETF 份额采集  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*55}")
        print(f"\n⏭️  昨天 {yesterday} 非A股交易日，今日无需运行，退出。")
        print(f"{'='*55}")
        return False
    return True


def get_next_trading_day(d):
    next_day = d + timedelta(days=1)
    while not is_workday(next_day):
        next_day += timedelta(days=1)
    return next_day


def is_last_trading_day_of_week(d):
    return get_next_trading_day(d).isocalendar()[1] != d.isocalendar()[1]


def is_last_trading_day_of_month(d):
    return get_next_trading_day(d).month != d.month


def is_last_trading_day_of_quarter(d):
    if d.month not in [3, 6, 9, 12]:
        return False
    return is_last_trading_day_of_month(d)


def get_market_value_target_row(yesterday):
    if is_last_trading_day_of_quarter(yesterday):
        print(f"   📊 昨天 {yesterday} 是季度最后交易日 → 市值填入第{ROW_QUARTERLY}行")
        return ROW_QUARTERLY
    elif is_last_trading_day_of_month(yesterday):
        print(f"   📊 昨天 {yesterday} 是月份最后交易日 → 市值填入第{ROW_MONTHLY}行")
        return ROW_MONTHLY
    elif is_last_trading_day_of_week(yesterday):
        print(f"   📊 昨天 {yesterday} 是本周最后交易日 → 市值填入第{ROW_WEEKLY}行")
        return ROW_WEEKLY
    else:
        print(f"   📊 昨天 {yesterday} 是普通交易日 → 市值填入第{ROW_DAILY}行")
        return ROW_DAILY


def shift_formula_rows(formula, delta):
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    def replace_row(match):
        return f"{match.group(1)}{int(match.group(2)) + delta}"
    return re.sub(r'([A-Z]+)(\d+)', replace_row, formula)


def prepare_excel_new_row():
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"❌ 找不到 Excel: {EXCEL_FILE_PATH}")
        return None

    try:
        with open(EXCEL_FILE_PATH, 'a'):
            pass
    except IOError:
        print(f"❌ 文件正在被占用！请先关闭 Excel 后再运行。")
        return None

    try:
        wb = load_workbook(EXCEL_FILE_PATH)
        ws = wb.active
        print(f"📋 工作表: {ws.title}")

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
        print(f"   找到 {len(code_to_col)} 个ETF代码映射")

        # 第一步：拍照第18行
        row18_snapshot = {
            col: ws.cell(DATA_START_ROW, col).value
            for col in range(1, ws.max_column + 1)
        }

        # 第二步：从下往上整块下移（MAX_HISTORY_ROW → 18，倒序）
        # 第18行→第19行，第19行→第20行……第163行→第164行
        # 第163行原内容会被覆盖（可接受），其余历史数据完整保留
        for row in range(MAX_HISTORY_ROW, DATA_START_ROW - 1, -1):
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row, col).value
                if isinstance(val, str) and val.startswith('='):
                    ws.cell(row + 1, col).value = shift_formula_rows(val, +1)
                else:
                    ws.cell(row + 1, col).value = val

        # 第三步：清空第18行
        for col in range(1, ws.max_column + 1):
            ws.cell(DATA_START_ROW, col).value = None

        # 第四步：只把公式原样写回第18行（数值保持空白，等待填入）
        for col, val in row18_snapshot.items():
            if isinstance(val, str) and val.startswith('='):
                ws.cell(DATA_START_ROW, col).value = val

        wb.save(EXCEL_FILE_PATH)
        print(f"   ✅ 第18~{MAX_HISTORY_ROW}行整块下移完成")
        print(f"   ✅ 第18行已清空并还原公式，等待数据填入")
        print("✅ 第一步完成")
        return code_to_col

    except PermissionError:
        print(f"❌ 文件写入被拒绝，请确认 Excel 已完全关闭后重试")
        return None
    except Exception as e:
        print(f"❌ Excel 准备失败: {e}")
        import traceback
        traceback.print_exc()
        return None


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
                    print(f"   ✅ 深交所 {code}: {shares_val/10000:.2f} 万份")
            else:
                print(f"   ⚠️ 深交所未找到: {code}")
    except Exception as e:
        print(f"❌ 深交所同步失败: {e}")
    return results_map


def extract_sse_data_dom(driver, code, max_retries=2):
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
                    raw_date  = tds[0].text.strip()
                    raw_share = tds[-1].text.replace(',', '').strip()
                    share_val = float(raw_share)

                    market_val = None
                    try:
                        wait.until(EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "div.js_transactionOverview td.colData_val")
                        ))
                        market_tds = driver.find_elements(
                            By.CSS_SELECTOR, "div.js_transactionOverview td.colData_val"
                        )
                        if market_tds:
                            raw_market = market_tds[0].text.replace(',', '').strip()
                            market_val = round(float(raw_market) / 10000, 6)
                    except:
                        pass

                    if share_val > 0:
                        parts = raw_date.split("-")
                        date_fmt = f"{parts[0]}//{int(parts[1])}/{int(parts[2])}"
                        return share_val, market_val, date_fmt

        except Exception as e:
            if attempt < max_retries:
                time.sleep(2)
            else:
                print(f"   ⚠️ DOM读取失败 {code}: {e}")
    return None, None, None


def fetch_sse_data(code_list):
    share_map     = {}
    market_map    = {}
    failed_list   = []
    detected_date = None

    options = Options()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
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
                share, market, date_str = extract_sse_data_dom(driver, code)
                if share:
                    share_map[str(code)] = share
                    if market is not None:
                        market_map[str(code)] = market
                    if detected_date is None and date_str:
                        detected_date = date_str
                    mkt_str = f"{market:.4f}亿" if market else "市值未获取"
                    print(f"   [{idx}/{len(code_list)}] ✅ 沪市 {code}: 份额={share:.2f}万份  市值={mkt_str}  日期={date_str}")
                else:
                    failed_list.append(code)
                    print(f"   [{idx}/{len(code_list)}] ⚠️ 沪市 {code} 失败")
            except Exception as e:
                failed_list.append(code)
                print(f"   [{idx}/{len(code_list)}] ❌ 沪市 {code} 异常: {e}")
    finally:
        driver.quit()

    if failed_list:
        print(f"\n⚠️ 未获取到数据的沪市ETF: {failed_list}")
    return share_map, market_map, detected_date


def fill_excel_data(code_to_col, share_dict, market_dict, data_date, market_target_row):
    try:
        wb = load_workbook(EXCEL_FILE_PATH)
        ws = wb.active

        if data_date:
            ws.cell(DATA_START_ROW, DATE_COL).value = data_date
            print(f"\n   📅 日期（来自上交所）: {data_date}")
        else:
            today = datetime.now()
            date_fmt = f"{today.year}//{today.month}/{today.day}"
            ws.cell(DATA_START_ROW, DATE_COL).value = date_fmt
            print(f"\n   📅 日期（系统备用）: {date_fmt}")

        filled_share = 0
        for code, share_val in share_dict.items():
            if code in code_to_col:
                share_col = code_to_col[code] + 2
                ws.cell(DATA_START_ROW, share_col).value = float(share_val)
                filled_share += 1
        print(f"   📈 份额填入完成: {filled_share}/{len(share_dict)} 只")

        filled_market = 0
        for code, market_val in market_dict.items():
            if code in code_to_col:
                market_col = code_to_col[code]
                ws.cell(market_target_row, market_col).value = float(market_val)
                filled_market += 1
                print(f"   💰 {code} 市值 {market_val:.4f}亿 → 第{market_target_row}行 col{market_col}")
        print(f"   💰 市值填入完成: {filled_market}/{len(market_dict)} 只 → 第{market_target_row}行")

        wb.save(EXCEL_FILE_PATH)
        print(f"\n✨ 全部完成！份额写入第{DATA_START_ROW}行，市值写入第{market_target_row}行")

    except Exception as e:
        print(f"❌ 数据填入失败（请先关闭文件）: {e}")
        import traceback
        traceback.print_exc()


def run_integration():
    if not check_trading_day():
        return

    yesterday = date.today() - timedelta(days=1)
    market_target_row = get_market_value_target_row(yesterday)

    start_time = datetime.now()
    print(f"{'='*55}")
    print(f"  ETF 份额采集  {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*55}")

    print("\n【第一步】准备 Excel 新行...")
    code_to_col = prepare_excel_new_row()
    if code_to_col is None:
        print("❌ Excel 准备失败，程序终止")
        return

    print("\n【第二步】抓取数据...")
    final_share = fetch_szse_data(SZSE_LIST)
    sse_share, market_dict, data_date = fetch_sse_data(SSE_LIST)
    final_share.update(sse_share)

    print("\n【第三步】填入数据...")
    fill_excel_data(code_to_col, final_share, market_dict, data_date, market_target_row)

    elapsed = (datetime.now() - start_time).seconds
    print(f"\n📊 采集汇总: 份额 {len(final_share)}/{len(SSE_LIST)+len(SZSE_LIST)} 只  市值 {len(market_dict)}/{len(SSE_LIST)} 只")
    print(f"⏱️  总耗时: {elapsed // 60} 分 {elapsed % 60} 秒")
    print(f"{'='*55}")


if __name__ == "__main__":
    run_integration()