from config import *

def run_step3_sync_to_kdocs():
    """第三阶段：同步到云端并关闭"""
    logger.info("[阶段3] 开始同步至金山文档...")
    df, app_xl, wb_xl = copy_local_excel_to_clipboard()

    options = Options()
    options.add_argument("--start-maximized")
    driver_k = webdriver.Chrome(options=options)
    wait_k = WebDriverWait(driver_k, 60)

    try:
        driver_k.get(KDOCS_URL)
        wait_k.until(EC.element_to_be_clickable((By.XPATH, "//div[text()='免费使用']"))).click()
        time.sleep(2)
        pyautogui.moveTo(317, 189, duration=0.5)
        pyautogui.click()
        time.sleep(5)
        wait_k.until(EC.element_to_be_clickable((By.ID, "confirmSync"))).click()
        time.sleep(5)

        search_box = wait_k.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.kdv-input__inner")))
        search_box.send_keys(SEARCH_NAME + Keys.ENTER)
        time.sleep(6)
        pyautogui.doubleClick(678, 412, duration=0.5)
        time.sleep(5)

        pyautogui.click(174, 1017)
        time.sleep(1)
        pyautogui.click(95, 302)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(5)
        logger.info("- 云端数据更新粘贴完成")

        clear_clipboard()
        logger.info("- 尝试关闭网页标签...")
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(2)
        pyautogui.click(735, 314)

        logger.info("- 正在关闭本地 Excel 进程...")
        wb_xl.close()
        app_xl.quit()
    except Exception as e:
        logger.error(f"同步出错: {e}")
        raise e
    finally:
        driver_k.quit()

if __name__ == "__main__":
    run_step3_sync_to_kdocs()