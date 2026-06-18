from config import *


def run_step2_process():
    """第二阶段：合并数据至本地Excel (纯数值模式)"""
    logger.info("[阶段2] 开始更新本地运营表...")
    force_close_excel(TARGET_FILE)
    wb = load_workbook(LOCAL_FILE)
    ws = wb[SHEET_NAME]

    def force_numeric(df, columns):
        for col in columns:
            if col in df.columns:
                # 1. 先转为字符串处理，防止类型报错
                df[col] = df[col].astype(str)

                # 2. 识别并处理带 % 的数值 (如 "15.5%" -> 0.155)
                mask = df[col].str.contains('%', na=False)
                # 使用 apply 处理以确保安全转换
                df.loc[mask, col] = df.loc[mask, col].str.replace('%', '', regex=False).astype(float) / 100

                # 3. 强制转为数字，无法转换的（如空值、其他杂质字符串）转为 0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df

    # 处理历史分析
    df_h = pd.read_excel(get_latest_file(DOWNLOAD_DIR, "历史分析"))
    df_h = force_numeric(df_h, ["近7日销量", "昨日销量"])
    write_dataframe(ws, df_h, ["SKU", "近7日销量", "昨日销量"], 5)

    # 处理商品分析
    df_p = pd.read_excel(get_latest_file(DOWNLOAD_DIR, "商品分析"))
    df_p = force_numeric(df_p, ["购物体验分", "折扣价(MXN)", "原价(MXN)", "近7天访问量"])
    write_dataframe(ws, df_p, ["商品ID", "购物体验分", "折扣价(MXN)", "原价(MXN)", "近7天访问量"], 10)

    # 处理广告
    f30, f7 = get_latest_ads_files(DOWNLOAD_DIR)

    # 广告 7天数据清洗 (包含 ACoS, TACOS 等百分比列)
    df_f7 = pd.read_excel(f7)
    df_f7 = force_numeric(df_f7,
                          ["展现量", "点击数", "ACoS", "广告花费(MXN)", "广告销量", "直接销量", "间接销量", "自然销量",
                           "TACOS(ACOAS)", "ROAS"])
    write_dataframe(ws, df_f7,
                    ["商品ID", "展现量", "点击数", "ACoS", "广告花费(MXN)", "广告销量", "直接销量", "间接销量",
                     "自然销量", "TACOS(ACOAS)", "ROAS"], 26)

    # 广告 30天数据清洗 (包含 TACOS 百分比列)
    df_f30 = pd.read_excel(f30)
    df_f30 = force_numeric(df_f30, ["TACOS(ACOAS)"])
    write_dataframe(ws, df_f30, ["商品ID", "TACOS(ACOAS)"], 38)

    wb.save(LOCAL_FILE)
    logger.info("- 本地文件更新保存成功 (已修复百分比识别问题)")


if __name__ == "__main__":
    run_step2_process()