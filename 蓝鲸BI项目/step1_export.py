from config import *

def run_step1_export(driver, wait):
    """第一阶段：导出数据"""
    logger.info("[阶段1] 开始后台导出数据...")
    driver.get(BASE_URL)
    wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="text"]'))).send_keys(USERNAME)
    driver.find_element(By.XPATH, '//input[@type="password"]').send_keys(PASSWORD)
    driver.find_element(By.XPATH, '//button').click()
    time.sleep(3)

    wait.until(EC.element_to_be_clickable((By.ID, "mercadoId"))).click()
    time.sleep(5)
    driver.execute_script("document.querySelectorAll('.el-dialog__wrapper, .v-modal').forEach(e => e.remove());")

    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'商品分析')]"))).click()
    time.sleep(6)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'导出明细')]"))).click()
    logger.info("- 商品分析导出已提交")

    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'交易分析')]"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'历史分析')]"))).click()
    time.sleep(5)
    select_time_range(driver, wait, "近7天")
    time.sleep(5)
    sku_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[.//span[contains(text(),'SKU每日表现')]]")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", sku_btn)
    time.sleep(3)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'导出明细')]"))).click()
    time.sleep(3)
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., '加入报表下载')]"))).click()
        logger.info("- 历史分析导出已提交")
    except:
        logger.warning("- 未发现'加入报表下载'按钮，跳过")
    driver.execute_script("document.querySelectorAll('.v-modal, .el-dialog__wrapper').forEach(e => e.style.display='none');")

    logger.info("进入广告流量，等待页面加载")
    ads_menu = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'广告流量')]")))
    driver.execute_script("arguments[0].click();", ads_menu)
    time.sleep(5)

    select_time_range(driver, wait, "近7天")
    export_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[.//span[contains(text(),'导出明细')]]")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", export_btn)
    time.sleep(1)
    ActionChains(driver).move_to_element(export_btn).pause(0.5).click().perform()
    time.sleep(1)

    try:
        msg_box = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, "el-message-box__wrapper")))
        confirm_btn = msg_box.find_element(By.XPATH, ".//button[contains(@class,'el-button--primary')]")
        driver.execute_script("arguments[0].click();", confirm_btn)
        logger.info("广告流量(7天)文件导出已提交")
    except Exception as e:
        logger.warning(f"未检测到广告(7天)确认弹窗: {e}")

    driver.execute_script("let masks = document.querySelectorAll('.v-modal'); masks.forEach(m => m.remove());")

    select_time_range(driver, wait, "近30天")
    export_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[.//span[contains(text(),'导出明细')]]")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", export_btn)
    time.sleep(1)
    ActionChains(driver).move_to_element(export_btn).pause(0.5).click().perform()
    time.sleep(1)

    try:
        msg_box = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, "el-message-box__wrapper")))
        confirm_btn = msg_box.find_element(By.XPATH, ".//button[contains(@class,'el-button--primary')]")
        driver.execute_script("arguments[0].click();", confirm_btn)
        logger.info("广告流量(30天)文件导出已提交")
    except Exception as e:
        logger.warning(f"未检测到广告(30天)确认弹窗: {e}")

    driver.execute_script("""
        let masks = document.querySelectorAll('.v-modal');
        masks.forEach(m => m.remove());
        let dialogs = document.querySelectorAll('.el-dialog__wrapper');
        dialogs.forEach(d => { if(d.style.display !== 'none') d.style.display = 'none'; });
    """)

    logger.info("等待480秒让文件生成并进入个人中心下载...")
    time.sleep(10)

    user_center = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'个人中心')]")))
    user_center.click()
    time.sleep(2)
    download_center = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'下载中心')]")))
    download_center.click()
    wait.until(EC.presence_of_element_located((By.XPATH, "//span[text()='下载']")))
    time.sleep(2)
    driver.execute_script("let masks = document.querySelectorAll('.v-modal'); masks.forEach(m => m.remove());")

    for i in range(1, 3):
        try:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, f"(//span[text()='下载'])[{i}]")))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(1)
            ActionChains(driver).move_to_element(btn).pause(0.3).click().perform()
            time.sleep(2)
        except Exception as e:
            logger.error(f"第{i}个下载失败: {e}")

    time.sleep(1)
    driver.quit()

if __name__ == "__main__":
    chrome_opts = Options()
    chrome_opts.add_argument("--disable-blink-features=AutomationControlled")
    main_driver = webdriver.Chrome(options=chrome_opts)
    main_driver.maximize_window()
    main_wait = WebDriverWait(main_driver, 60)
    run_step1_export(main_driver, main_wait)