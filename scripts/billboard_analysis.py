#!/usr/bin/env python3
"""
🏆 全市场龙虎榜基础数据 + 技术面指标（GitHub Actions 自动化版）
====================================================
修改说明：
  - from jqdata import * → from jqdatasdk import * + auth()
  - IPython.display → print（console 兼容）
  - 支持输出 JSON / CSV（适合博客读取）
  - 输出路径可配置，默认 output/ 目录
  - 自动判断最新交易日

使用方式：
  1. 在下方 "填写区" 填入聚宽账号密码
  2. python billboard_analysis.py

环境变量（可选，优先级低于填写区）：
  JQ_USER        聚宽账号（手机号或邮箱）
  JQ_PASS        聚宽密码
  OUTPUT_FORMAT  输出格式 json/csv/both（默认 both）
  OUTPUT_DIR     输出目录（默认 output）
"""

# =============================================================================
# 🔑 填写区 —— 在这里填入你的聚宽账号密码
# =============================================================================
JQ_USER = ''      # ← 填你的聚宽账号（手机号或邮箱）
JQ_PASS = ''      # ← 填你的聚宽密码
# =============================================================================

import os
import sys
import json
import pandas as pd
import numpy as np
from jqdatasdk import *
import datetime
import math

# =============================================================================
# 配置区
# =============================================================================
OUTPUT_DIR = os.environ.get('OUTPUT_DIR', 'output')
OUTPUT_FORMAT = os.environ.get('OUTPUT_FORMAT', 'both')  # json / csv / both

# =============================================================================
# 🛠️ 核心一级火箭：基础动能、去重吞噬与数板引擎
# =============================================================================

def calculate_continuous_boards(code, end_date_str):
    """连续涨跌停阶梯计数器（智能识别个股板幅）"""
    try:
        # 智能识别是否为创业板或科创板（20%板幅）
        is_20pct = code.startswith(('30', '68'))

        end_dt = datetime.datetime.strptime(end_date_str, '%Y-%m-%d').date()
        raw_start = end_dt - datetime.timedelta(days=60)
        hist_days = get_trade_days(start_date=raw_start, end_date=end_dt)
        if len(hist_days) == 0:
            return '0'

        prices = get_price(code, start_date=hist_days[0], end_date=end_dt,
                          frequency='daily', fields=['close', 'pre_close'])
        if prices.empty:
            return '0'

        prices = prices.sort_index(ascending=False)
        limit_ratio = 0.199 if is_20pct else 0.099

        up_count, down_count = 0, 0
        state = 'INIT'

        for idx, row in prices.iterrows():
            close_p = row['close']
            pre_close = row['pre_close']
            if pre_close <= 0:
                break
            pct = (close_p - pre_close) / pre_close

            is_up_limit = (pct >= limit_ratio)
            is_down_limit = (pct <= -limit_ratio)

            if state == 'INIT':
                if is_up_limit:
                    up_count += 1
                    state = 'UP'
                elif is_down_limit:
                    down_count += 1
                    state = 'DOWN'
                else:
                    break
            elif state == 'UP':
                if is_up_limit:
                    up_count += 1
                else:
                    break
            elif state == 'DOWN':
                if is_down_limit:
                    down_count += 1
                else:
                    break

        if up_count > 0:
            return f"涨{up_count}"
        if down_count > 0:
            return f"跌{down_count}"
        return '0'
    except Exception:
        return '0'


# =============================================================================
# 📈 核心技术面量化计算引擎（ATH/ATL、1年/半年位置与最大回撤、均线与均量）
# =============================================================================

def calculate_technical_features(code, target_date_str):
    """
    计算特定个股在特定日期的所有技术面指标：
    1. 周线级别 ATH/ATL
    2. 1年内价格定位、偏离度及年内历史最大回撤（250交易日）
    3. 半年内价格定位及偏离度（125交易日）
    4. 日线级别 60日均量突破、均线健康度判定
    """
    tech_results = {
        'ATH/ATL': '否',
        '一年内价格': '年中价',
        '距年高回撤(%)': np.nan,
        '距年低反弹(%)': np.nan,
        '年内最大回撤(%)': np.nan,
        '半年内价格': '半年中价',
        '距半年高回撤(%)': np.nan,
        '距半年低反弹(%)': np.nan,
        '放量MA60': '否',
        '均线健康否': '否'
    }

    try:
        target_dt = datetime.datetime.strptime(target_date_str, '%Y-%m-%d').date()

        # --- 1. 日线级别数据获取（获取 350 天日线确保计算 MA250 和年线极值） ---
        trade_days_daily = get_trade_days(end_date=target_dt, count=350)
        if len(trade_days_daily) < 250:
            return tech_results

        daily_df = get_price(code, start_date=trade_days_daily[0], end_date=target_dt,
                            frequency='daily', fields=['close', 'volume'])
        if daily_df.empty or len(daily_df) < 250:
            return tech_results

        close_series = daily_df['close']
        volume_series = daily_df['volume']

        current_volume = volume_series.iloc[-1]
        current_close = close_series.iloc[-1]

        # 计算 60日均量
        ma60_volume = volume_series.iloc[-60:].mean()
        tech_results['放量MA60'] = '是' if current_volume > ma60_volume else '否'

        # 计算 10, 20, 60, 250 日均线
        ma10 = close_series.iloc[-10:].mean()
        ma20 = close_series.iloc[-20:].mean()
        ma60 = close_series.iloc[-60:].mean()
        ma250 = close_series.iloc[-250:].mean()

        # 判断均线多头排列（均线健康否）
        tech_results['均线健康否'] = '是' if (ma10 > ma20 > ma60 > ma250) else '否'

        # 1.2 过去一年（250个交易日）的价格判断与【年内最大回撤】计算
        past_year_closes = close_series.iloc[-250:]
        year_max = past_year_closes.max()
        year_min = past_year_closes.min()

        # 1.2.1 一年内价格定位
        if current_close >= year_max:
            tech_results['一年内价格'] = '年高价'
        elif current_close <= year_min:
            tech_results['一年内价格'] = '年低价'
        else:
            tech_results['一年内价格'] = '年中价'
        tech_results['距年高回撤(%)'] = ((year_max - current_close) / year_max) * 100
        tech_results['距年低反弹(%)'] = ((current_close - year_min) / year_min) * 100

        # 1.2.2 过去250个交易日的历史最大回撤 (MDD)
        roll_max = past_year_closes.cummax()
        drawdowns = (roll_max - past_year_closes) / roll_max
        tech_results['年内最大回撤(%)'] = drawdowns.max() * 100

        # 1.3 过去半年（125个交易日）的价格判断
        past_half_year_closes = close_series.iloc[-125:]
        half_year_max = past_half_year_closes.max()
        half_year_min = past_half_year_closes.min()

        if current_close >= half_year_max:
            tech_results['半年内价格'] = '半年高价'
        elif current_close <= half_year_min:
            tech_results['半年内价格'] = '半年低价'
        else:
            tech_results['半年内价格'] = '半年中价'
        tech_results['距半年高回撤(%)'] = ((half_year_max - current_close) / half_year_max) * 100
        tech_results['距半年低反弹(%)'] = ((current_close - half_year_min) / half_year_min) * 100

        # --- 2. 周线级别数据获取（ATH / ATL 计算） ---
        start_history = target_dt - datetime.timedelta(days=365 * 10)
        week_df = get_price(code, start_date=start_history, end_date=target_dt,
                           frequency='weekly', fields=['close'])

        if not week_df.empty:
            current_week_close = week_df['close'].iloc[-1]
            all_time_max = week_df['close'].max()
            all_time_min = week_df['close'].min()

            if current_week_close >= all_time_max:
                tech_results['ATH/ATL'] = 'ATH'
            elif current_week_close <= all_time_min:
                tech_results['ATH/ATL'] = 'ATL'
            else:
                tech_results['ATH/ATL'] = '否'

    except Exception:
        pass

    return tech_results


# =============================================================================
# 🚀 核心二级火箭：龙虎榜股票全景数据诊断引擎（精简锚点版）
# =============================================================================
def run_billboard_stocks_diagnostic_engine(anchor_date_str=None, data_days=255):
    """
    V15.2 龙虎榜股票靶向多维数据提取与分市场归类诊断引擎
    :param anchor_date_str: 🌟 锚点基准日
    :param data_days: 🌟 数据回溯天数
    """
    print(f"\n{'='*77}")
    print(f"🎬 正在初始化龙虎榜股票全景提取与多维分市场归类诊断引擎 (V15.2 锚点精简版)...")
    print(f"{'='*77}")

    # 1. 确定锚点基准日
    start_bound = datetime.date.today() - datetime.timedelta(days=730)
    all_trade_days = get_trade_days(start_date=start_bound, end_date=datetime.date.today())

    if anchor_date_str is None:
        anchor_date = all_trade_days[-1]
    else:
        anchor_date = pd.to_datetime(anchor_date_str).date()

    bb_str = anchor_date.strftime('%Y-%m-%d')
    print(f">>> 控制状态 -> 锚点基准日 (提取当日龙虎榜): {bb_str} | 技术面数据回溯天数: {data_days} 个交易日")

    # --- 🎯 2. 直接提取锚点当天的龙虎榜股票名单 ---
    print(f">>> 步骤一：正在检索锚点基准日 [{bb_str}] 当天的龙虎榜单...")
    try:
        raw_bb_df = get_billboard_list(stock_list=None, start_date=bb_str, end_date=bb_str)
        if raw_bb_df.empty:
            print(f"⚠️ 提示：在锚点基准日 {bb_str} 期间没有任何股票上榜。")
            return pd.DataFrame()

        filtered_bb = raw_bb_df[raw_bb_df['code'].str.startswith(('00', '30', '60', '68', '43', '83', '87', '88'))].copy()
        target_codes = filtered_bb['code'].unique().tolist()
        print(f">>> 🏆 成功在龙虎榜中捕获到 {len(target_codes)} 只异动明星股！")
    except Exception as e:
        print(f"❌ 错误：提取龙虎榜股票失败: {e}")
        return pd.DataFrame()

    if not target_codes:
        print("⚠️ 当日无龙虎榜数据，跳过运行。")
        return pd.DataFrame()

    # --- 🎯 3. 计算提取个股基础数据的时间轴 ---
    print(f">>> 步骤二：正在配置基础数据时间轴...")
    data_trade_days = get_trade_days(end_date=anchor_date, count=data_days)
    if len(data_trade_days) == 0:
        print("❌ 错误：未获取到数据回溯交易日。")
        return pd.DataFrame()

    start_date_str = data_trade_days[0].strftime('%Y-%m-%d')
    end_date_str = data_trade_days[-1].strftime('%Y-%m-%d')
    print(f">>> 数据回溯时间轴：{start_date_str} 至 {end_date_str} (共 {len(data_trade_days)} 个交易日)")

    # --- 4. 内存映射：获取申万三级行业及成分股列表 ---
    print(">>> 正在初始化全市场【申万三级】个股归属矩阵及股票池...")
    sw3_mapping = {}
    industry_to_stocks = {}
    all_unique_stocks = set()
    try:
        sw3_industries = get_industries(name='sw_l3')
        for ind_code, row in sw3_industries.iterrows():
            ind_name = row['name']
            stocks_in_ind = get_industry_stocks(ind_code, date=anchor_date)
            if stocks_in_ind:
                industry_to_stocks[ind_name] = stocks_in_ind
                all_unique_stocks.update(stocks_in_ind)
                for s in stocks_in_ind:
                    sw3_mapping[s] = ind_name
    except Exception as e:
        print(f"⚠️ 申万三级列表构建轻微扰动: {e}")

    # --- 5. 核心合成：成分股等权合成行业每日涨跌幅 ---
    print(">>> 正在通过【成分股等权合成法】计算相关三级行业每日真实涨跌幅矩阵...")
    sw3_daily_perf = {}
    try:
        all_unique_stocks = list(all_unique_stocks)
        if all_unique_stocks:
            extended_trade_days = get_trade_days(end_date=anchor_date, count=len(data_trade_days) + 1)
            ext_start_str = extended_trade_days[0].strftime('%Y-%m-%d')

            stock_prices = get_price(all_unique_stocks, start_date=ext_start_str, end_date=end_date_str,
                                    frequency='daily', fields=['close'], panel=False)

            if not stock_prices.empty:
                stock_prices['time'] = pd.to_datetime(stock_prices['time']).dt.strftime('%Y-%m-%d')
                price_pivot = stock_prices.pivot(index='time', columns='code', values='close')
                stock_pct_matrix = price_pivot.pct_change() * 100

                target_dates = [d.strftime('%Y-%m-%d') for d in data_trade_days]
                stock_pct_matrix = stock_pct_matrix.loc[stock_pct_matrix.index.isin(target_dates)]

                for date_str in stock_pct_matrix.index:
                    sw3_daily_perf[date_str] = {}
                    day_stock_perf = stock_pct_matrix.loc[date_str]
                    for ind_name, stocks in industry_to_stocks.items():
                        ind_stock_pcts = day_stock_perf.reindex(stocks).dropna()
                        if not ind_stock_pcts.empty:
                            sw3_daily_perf[date_str][ind_name] = ind_stock_pcts.mean()
                print(">>> 🛠️ 行业涨跌幅矩阵内存合成完毕！")
            else:
                print(">>> ⚠️ 行业数据为空，跳过行业涨跌幅计算")
    except Exception as e:
        print(f"⚠️ 行业涨跌幅合成异常: {e}")

    # --- 6. 智能多维度指数对齐准备 ---
    index_codes = ['000001.XSHG', '399001.XSHE', '399102.XSHE']
    index_perf_dict = {}
    try:
        index_prices = get_price(index_codes, start_date=start_date_str, end_date=end_date_str,
                                frequency='daily', fields=['close', 'pre_close'], panel=False)
        index_prices['time'] = pd.to_datetime(index_prices['time']).dt.strftime('%Y-%m-%d')
        index_prices['pct'] = (index_prices['close'] - index_prices['pre_close']) / index_prices['pre_close'] * 100
        index_perf_dict = index_prices.groupby('time').apply(
            lambda x: x.set_index('code')['pct'].to_dict()
        ).to_dict()
    except Exception as e:
        print(f"⚠️ 指数数据获取异常: {e}")

    all_results = []
    total_days = len(data_trade_days)

    # --- 7. 纵向时间轴与个股穿透计算 ---
    print(">>> 正在纵向穿透个股的多维财务与技术面诊断指标...")
    for day_idx, current_date in enumerate(data_trade_days):
        dt_str = current_date.strftime('%Y-%m-%d')

        if day_idx % 50 == 0:
            print(f"    → 进度: {day_idx}/{total_days} 天 ({dt_str})")

        # 批量拉取行情数据
        try:
            daily_price = get_price(target_codes, start_date=dt_str, end_date=dt_str,
                                   frequency='daily',
                                   fields=['open', 'close', 'high', 'low', 'money', 'volume', 'pre_close'],
                                   panel=False)
            if daily_price.empty:
                continue

            if 'code' not in daily_price.columns:
                daily_price = daily_price.reset_index()

            daily_price['change_pct'] = (daily_price['close'] - daily_price['pre_close']) / daily_price['pre_close'] * 100
            daily_price['calc_amp'] = (daily_price['high'] - daily_price['low']) / daily_price['pre_close'] * 100
            price_dict = daily_price.set_index('code').to_dict(orient='index')
        except Exception:
            continue

        # 抓取基本面财务数据
        fund_dict = {}
        try:
            q = query(
                valuation.code, valuation.market_cap, valuation.pe_ratio, valuation.turnover_ratio,
                indicator.roe, indicator.eps,
                indicator.inc_total_revenue_year_on_year,
                indicator.inc_net_profit_year_on_year
            ).filter(valuation.code.in_(target_codes))
            fund_df = get_fundamentals(q, date=dt_str)
            if not fund_df.empty:
                fund_dict = fund_df.set_index('code').to_dict(orient='index')
        except Exception:
            pass

        # 查询去年同期的 EPS
        last_year_dict = {}
        try:
            last_year_date = current_date - datetime.timedelta(days=365)
            last_year_str = last_year_date.strftime('%Y-%m-%d')
            q_ly = query(indicator.code, indicator.eps).filter(indicator.code.in_(target_codes))
            ly_fund_df = get_fundamentals(q_ly, date=last_year_str)
            if not ly_fund_df.empty:
                last_year_dict = ly_fund_df.set_index('code')['eps'].to_dict()
        except Exception:
            pass

        # 精准对照当天的具体上榜原因
        today_bb_lbls = {}
        try:
            bb_df = get_billboard_list(stock_list=target_codes, start_date=dt_str, end_date=dt_str)
            if not bb_df.empty:
                today_bb_lbls = bb_df.set_index('code')['abnormal_name'].to_dict()
        except Exception:
            pass

        day_raw_records = []
        for code in target_codes:
            if code not in price_dict:
                continue

            p_info = price_dict.get(code, {})
            f_info = fund_dict.get(code, {})

            # 🔍 计算高级技术面数据
            tech_info = calculate_technical_features(code, dt_str)

            # 智能判定偏离值锚定指数
            if code.endswith('.XSHG'):
                idx_key = '000001.XSHG'
            elif code.startswith('30'):
                idx_key = '399102.XSHE'
            else:
                idx_key = '399001.XSHE'

            base_index_pct = index_perf_dict.get(dt_str, {}).get(idx_key, 0.0)

            close_price_val = p_info.get('close', np.nan)
            stock_pct = p_info.get('change_pct', 0.0)
            turnover_val = f_info.get('turnover_ratio', 0.0)
            amplitude_val = p_info.get('calc_amp', 0.0)
            deviation_val = stock_pct - base_index_pct

            industry_name = sw3_mapping.get(code, "未知三级行业")
            industry_pct_val = sw3_daily_perf.get(dt_str, {}).get(industry_name, np.nan)

            current_eps = f_info.get('eps', np.nan)
            last_year_eps = last_year_dict.get(code, np.nan)

            if pd.notna(current_eps) and pd.notna(last_year_eps) and last_year_eps != 0:
                calc_eps_yoy = ((current_eps - last_year_eps) / abs(last_year_eps)) * 100
            else:
                calc_eps_yoy = np.nan

            abnormal_name_str = today_bb_lbls.get(code, '未于当日上榜(区间股)')
            raw_volume = p_info.get('volume', 0.0)
            volume_in_hands = raw_volume / 100.0 if pd.notna(raw_volume) else 0.0

            # 将 None 替换为 NaN 以保证 JSON 序列化
            def safe_val(v):
                return v if v is not None else np.nan

            record = {
                '日期': dt_str,
                '今日上榜原因': abnormal_name_str,
                '股票代码': code,
                '股票名称': get_security_info(code).display_name if get_security_info(code) else '未知',
                '连板标识': calculate_continuous_boards(code, dt_str),
                '收盘价': safe_val(close_price_val),
                '涨跌幅': stock_pct,

                # --- 🔍 高级技术面诊断指标 ---
                'ATH/ATL': tech_info['ATH/ATL'],

                # 一年周期维度
                '一年内价格': tech_info['一年内价格'],
                '距年高回撤(%)': safe_val(tech_info['距年高回撤(%)']),
                '距年低反弹(%)': safe_val(tech_info['距年低反弹(%)']),
                '年内最大回撤(%)': safe_val(tech_info['年内最大回撤(%)']),

                # 半年周期维度
                '半年内价格': tech_info['半年内价格'],
                '距半年高回撤(%)': safe_val(tech_info['距半年高回撤(%)']),
                '距半年低反弹(%)': safe_val(tech_info['距半年低反弹(%)']),

                # 量能与均线健康度
                '成交量突破(MA60)': tech_info['放量MA60'],
                '均线健康否': tech_info['均线健康否'],
                '当日成交量(手)': volume_in_hands,

                # --- 基本面财务数据及龙虎榜统计 ---
                '换手率(%)': turnover_val,
                '振幅(%)': amplitude_val,
                '偏离值(%)': deviation_val,
                '一年上榜次数': 0,
                '总市值': safe_val(f_info.get('market_cap', np.nan)),
                '当日交易额(亿)': safe_val(p_info.get('money', 0.0) / 100000000.0),
                'EPS': current_eps,
                'PE(TTM)': safe_val(f_info.get('pe_ratio', np.nan)),
                'ROE(%)': safe_val(f_info.get('roe', np.nan)),
                '销售额同比(%)': safe_val(f_info.get('inc_total_revenue_year_on_year', np.nan)),
                '净利润同比(%)': safe_val(f_info.get('inc_net_profit_year_on_year', np.nan)),
                'EPS同比(%)': calc_eps_yoy,
                '所属行业': industry_name,
                '行业涨跌': industry_pct_val
            }
            day_raw_records.append(record)

        if not day_raw_records:
            continue
        all_results.extend(day_raw_records)

    if not all_results:
        print("⚠️ 提示：所选数据时间段内未捕获到任何个股的有效行情数据。")
        return pd.DataFrame()

    df_final = pd.DataFrame(all_results)

    # 追溯近一年该股票被龙虎榜统计的总次数
    try:
        year_start = anchor_date - datetime.timedelta(days=365)
        year_bb = get_billboard_list(
            stock_list=target_codes,
            start_date=year_start.strftime('%Y-%m-%d'),
            end_date=anchor_date.strftime('%Y-%m-%d')
        )
        if not year_bb.empty:
            freq_series = year_bb.groupby('code').apply(
                lambda x: len(x[['day', 'abnormal_name']].drop_duplicates())
            )
            df_final['一年上榜次数'] = df_final['股票代码'].map(freq_series).fillna(0).astype(int)
    except Exception:
        pass

    # =============================================================================
    # 🎯 核心升级：中国A股多层次资本市场分级排布与智能归类
    # =============================================================================
    def get_market_rank(code):
        """根据股票代码前缀进行分市场等级计算"""
        if code.startswith('60'):
            return 1  # 沪市主板
        elif code.startswith('00'):
            return 2  # 深市主板
        elif code.startswith('30'):
            return 3  # 创业板
        elif code.startswith('68'):
            return 4  # 科创板
        elif code.startswith(('8', '4', '43', '83', '87', '88')) or code.endswith('.BJ'):
            return 5  # 北交所
        else:
            return 6  # 其他

    df_final['_market_rank'] = df_final['股票代码'].apply(get_market_rank)

    exact_col_order = [
        '日期', '股票代码', '股票名称', '连板标识', '今日上榜原因', '收盘价', '涨跌幅',
        'ATH/ATL',
        '一年内价格', '距年高回撤(%)', '距年低反弹(%)', '年内最大回撤(%)',
        '半年内价格', '距半年高回撤(%)', '距半年低反弹(%)',
        '成交量突破(MA60)', '均线健康否', '当日成交量(手)',
        '换手率(%)', '振幅(%)', '偏离值(%)', '一年上榜次数',
        '总市值', '当日交易额(亿)',
        'EPS', 'PE(TTM)', 'ROE(%)', '销售额同比(%)', '净利润同比(%)', 'EPS同比(%)',
        '所属行业', '行业涨跌'
    ]
    exact_col_order = [c for c in exact_col_order if c in df_final.columns]

    # 🚀 排序优先级：日期倒序 -> 市场分级正序 -> 代码升序
    df_final = df_final.sort_values(
        by=['日期', '_market_rank', '股票代码'],
        ascending=[False, True, True]
    ).reset_index(drop=True)

    df_final.drop(columns=['_market_rank'], inplace=True, errors='ignore')
    df_final = df_final[exact_col_order]

    # Pandas 格式设置
    pd.set_option('display.float_format', '{:.2f}'.format)
    pd.set_option('display.max_columns', 35)

    if '当日成交量(手)' in df_final.columns:
        df_final['当日成交量(手)'] = df_final['当日成交量(手)'].astype(int)

    print(f"\n>>> 🎯 提取完成！数据行数: {len(df_final)}")

    return df_final


# =============================================================================
# 🎛️ 主入口
# =============================================================================
def main():
    # 0. 认证（优先使用填写区的值，其次读取环境变量）
    jq_user = JQ_USER or os.environ.get('JQ_USER')
    jq_pass = JQ_PASS or os.environ.get('JQ_PASS')
    if not jq_user or not jq_pass:
        print("❌ 错误：请在上方「填写区」填入聚宽账号密码，或设置环境变量 JQ_USER / JQ_PASS")
        sys.exit(1)

    print(">>> 正在认证聚宽数据服务...")
    auth(jq_user, jq_pass)
    print(">>> ✅ 认证成功！")

    # 1. 创建输出目录
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 2. 运行诊断引擎
    anchor_date_str = None  # None = 最新交易日
    df = run_billboard_stocks_diagnostic_engine(
        anchor_date_str=anchor_date_str,
        data_days=255
    )

    if df.empty:
        print("⚠️ 没有数据输出，跳过保存。")
        return

    # 3. 显示前几行
    print("\n>>> 数据预览（前5行）：")
    print(df.head().to_string())

    # 4. 生成日期戳
    today_str = datetime.date.today().strftime('%Y%m%d')

    # 5. 输出 JSON（适合博客前端读取）
    if OUTPUT_FORMAT in ('json', 'both'):
        json_path = os.path.join(OUTPUT_DIR, f'billboard_data_{today_str}.json')
        # 将 NaN 转成 null 以便 JSON 序列化
        df_json = df.replace({np.nan: None})
        records = df_json.to_dict(orient='records')
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        print(f"📄 JSON 输出: {json_path} ({len(records)} 条记录)")

        # 同时输出 latest.json（始终指向最新，方便博客引用固定路径）
        latest_json = os.path.join(OUTPUT_DIR, 'billboard_data_latest.json')
        with open(latest_json, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        print(f"📄 最新快照: {latest_json}")

    # 6. 输出 CSV
    if OUTPUT_FORMAT in ('csv', 'both'):
        csv_path = os.path.join(OUTPUT_DIR, f'billboard_data_{today_str}.csv')
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"📄 CSV 输出: {csv_path}")

        latest_csv = os.path.join(OUTPUT_DIR, 'billboard_data_latest.csv')
        df.to_csv(latest_csv, index=False, encoding='utf-8-sig')
        print(f"📄 最新快照: {latest_csv}")

    print(f"\n✅ 全部完成！输出目录: {OUTPUT_DIR}")


if __name__ == '__main__':
    main()
