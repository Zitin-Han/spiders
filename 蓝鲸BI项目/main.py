import sys
from config import logger, TARGET_FILE, force_close_excel
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait

# 导入拆分后的模块
import step1_export
import step2_process
import step3_sync

if __name__ == "__main__":
    try:
        logger.info("========== 开始执行任务 (单次运行模式) ==========")

        # 1. 环境清理：强制关闭可能占用 Excel 的进程
        force_close_excel(TARGET_FILE)

        # 2. 执行第一阶段：后台导出数据
        chrome_opts = Options()
        chrome_opts.add_argument("--disable-blink-features=AutomationControlled")
        main_driver = webdriver.Chrome(options=chrome_opts)
        main_driver.maximize_window()
        main_wait = WebDriverWait(main_driver, 60)

        step1_export.run_step1_export(main_driver, main_wait)
        # 注意：step1_export 内部已包含 main_driver.quit()

        # 3. 执行第二阶段：本地数据处理与写入
        step2_process.run_step2_process()

        # 4. 执行第三阶段：同步至金山文档
        step3_sync.run_step3_sync_to_kdocs()

        logger.info("========== 全流程自动化执行成功 ==========")

    except Exception as e:
        logger.error(f"程序运行出错: {e}")

        # 异常情况下尝试关闭浏览器
        try:
            if 'main_driver' in locals():
                main_driver.quit()
        except:
            pass

        sys.exit(1)  # 发生错误时以状态码 1 退出