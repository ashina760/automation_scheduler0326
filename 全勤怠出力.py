import os
import shutil
import sys
import pandas as pd
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from webdriver_manager.chrome import ChromeDriverManager
from FlaskTaskQueue import download_folder
from FlaskTaskQueue import destination_folder

sys.stdout.flush()
download_folder = download_folder
destination_folder = destination_folder
subfolder_name = "全勤怠"
destination_folder = os.path.join(destination_folder, subfolder_name)
os.makedirs(destination_folder, exist_ok=True)

def main():
    try:
        sys.stdout.flush()
        options = Options()
        # 创建 ChromeDriver 实例
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        print("web open", flush=True)
        # 打开网页
        driver.get("https://ccbji-cjk.company.works-hi.com/CPNYCJK/cjkweb")
        driver.maximize_window()

        # 登录操作
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "txt_login_user"))).send_keys("CCBJI_KT2019")
        driver.find_element(By.ID, "txt_login_pass").send_keys("#ccbji2019kt")
        driver.find_element(By.XPATH, "//span[text()='ログイン']").click()

        # 点击"お気に入り"
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "favoriteAll"))).click()
        
        # 点击 dt 元素
        dt_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'dl#srw_arj_manager dt')))
        dt_element.click()

        # 切换到新标签页
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
        driver.switch_to.window(driver.window_handles[1])

        # 双击"一般"元素
        sleep(10)
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[@class='ctv-caption' and text()='一般']")))
        actions = ActionChains(driver)
        actions.double_click(element).perform()

        # 执行后续操作
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/table/tbody/tr[3]/td[1]/div/div/table/tbody/tr[3]/td/div/div[2]/div/ul/li[1]/ul/li[11]/span[3]")))
        driver.execute_script("arguments[0].click();", element)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[@class='ui-button-text' and text()='ジョブの実行']"))).click()
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='cpnyDialogButton1']/span"))).click()

        # 关闭当前标签页，返回主页面
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        print("wait 0.5H", flush=True)
        sleep(1)
        # 重复相同操作
        dt_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'dl#srw_arj_manager dt')))
        dt_element.click()

        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
        driver.switch_to.window(driver.window_handles[1])

        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[@class='ctv-caption' and text()='一般']")))
        actions.double_click(element).perform()

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='arj_grp_def_tree']/div/ul/li[1]/ul/li[11]/span[3]"))).click()

        driver.execute_script("arguments[0].click();", WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//label[@class='sheet-content' and text()='タスク1']"))))

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='toolbtn_csv_out_1']/span[2]"))).click()

        print("csv download", flush=True)
        sleep(120)
        
        # 处理下载的 CSV 文件
        
        files = [os.path.join(download_folder, f) for f in os.listdir(download_folder)]
        latest_file = max(files, key=os.path.getctime)
        shutil.move(latest_file, destination_folder)

        # 读取 CSV，保存为 xlsb 格式
        csv_file = os.path.join(destination_folder, os.path.basename(latest_file))
        df = pd.read_csv(csv_file, encoding='CP932', low_memory=False)
        print("change csv to xlsb", flush=True)
        with xw.App(visible=False) as app:
            wb = app.books.add()
            ws = wb.sheets[0]
            ws.range("A1").value = df.values
            ws.range("A1").value = df.columns.tolist()
            from datetime import datetime
            today = datetime.today()
            file_name = f'{today.strftime('%Y%m')}全勤怠DB_{today.strftime('%Y%m%d')}.xlsb'
            xlsb_file = os.path.join(destination_folder, file_name)
            wb.save(xlsb_file)
            wb.close()

        print(f"save as {xlsb_file}")
        print(f"DESTINATION_FOLDER: {destination_folder}")
        print("finish")
        sys.exit(0)

    except Exception as e:
        print(f"err: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()

