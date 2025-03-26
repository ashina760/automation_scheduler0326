import base64
import os
import sys
import json
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from webdriver_manager.chrome import ChromeDriverManager
from FlaskTaskQueue import download_folder
from FlaskTaskQueue import destination_folder

sys.stdout.flush()

CONFIG_PATH = None
with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = json.load(f)
userid = config["userid"]
userpw = config["userpw"]
url = None
download_folder = download_folder
destination_folder = destination_folder
subfolder_name = "全区分集計連携"
destination_folder = os.path.join(destination_folder, subfolder_name)
os.makedirs(destination_folder, exist_ok=True)


def main():
    try:
        sys.stdout.flush()
        # 初始化 WebDriver
        service = Service()
        options = Options()
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("open web")
        # 打开网站
        driver.get(url)
        driver.maximize_window()
        wait = WebDriverWait(driver, 30)

        # 登录操作
        user_input = wait.until(EC.presence_of_element_located((By.ID, "txt_login_user")))
        user_input.send_keys("userid")

        password_input = wait.until(EC.presence_of_element_located((By.ID, "txt_login_pass")))
        password_input.send_keys("#userpw")

        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='ログイン']")))
        login_button.click()

        # 点击 "お気に入り" 链接
        favorite_link = wait.until(EC.element_to_be_clickable((By.ID, "favoriteAll")))
        favorite_link.click()

        # 选择 dt 元素
        dt_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "dl#srw_dl_setting dt")))
        dt_element.click()

        # 切换到第二个标签页
        wait.until(lambda d: len(d.window_handles) > 1)
        driver.switch_to.window(driver.window_handles[1])

        # 点击目标 div
        target_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='cg-col0' and text()='SYS_MST_UPDATE04']")))
        target_element.click()

        # 点击执行按钮
        execute_button = wait.until(EC.element_to_be_clickable((By.ID, "toolbtn_execute")))
        execute_button.click()

        # 取消操作
        # cancel_button = wait.until(EC.element_to_be_clickable((By.ID, "btn_common_dialog_ok")))
        cancel_button = wait.until(EC.element_to_be_clickable((By.ID, "btn_common_dialog_cancel")))
        cancel_button.click()

        sleep(10)
        # 关闭当前标签并返回主窗口
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

        # 点击 batch_exec dt
        batch_exec = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "dl#batch_exec dt")))
        batch_exec.click()

        # 切换到新标签页
        wait.until(lambda d: len(d.window_handles) > 1)
        driver.switch_to.window(driver.window_handles[1])

        # # 选择下拉框值
        # select_element = wait.until(EC.presence_of_element_located((By.ID, "EcecGrpCombo")))
        # select = Select(select_element)
        # select.select_by_value("40")

        # 点击复选框
        checkbox_600 = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[td/div[text()='600']]//input[@type='checkbox']")))
        checkbox_600.click()
        sleep(10)
        # 点击 601 对应的复选框
        checkbox_601 = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[td/div[text()='601']]//input[@type='checkbox']")))
        checkbox_601.click()

        # 点击 "最初から全て実行" 按钮
        batch_execute_btn = wait.until(EC.element_to_be_clickable((By.ID, "BatchExecuteBtn")))
        batch_execute_btn.click()

        # 点击 "キャンセル" 按钮
        cancel_button_2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='キャンセル']")))
        cancel_button_2.click()

        sleep(10)
        print("wati 3H")
        # 找到所有 class 为 cg-col3 的 div 元素

        # while True:

        table = driver.find_element(By.ID, 'ExecPgmGrid_tbody')
        # 在 table 元素中查找所有 class 为 cg-col3 的 div 元素
        elements = table.find_elements(By.CLASS_NAME, 'cg-col3')

        if all(element.text == '正常終了' for element in elements):
            print_options = {
            'landscape': True,  # 横向
            'format': 'A4',  # A4纸
            'printBackground': True,  # 打印背景
            'marginTop': 0,  # 顶部边距
            'marginBottom': 0,  # 底部边距
            'marginLeft': 0,  # 左边距
            'marginRight': 0  # 右边距
        }

            # 生成 PDF 并保存
            pdf = driver.execute_cdp_cmd('Page.printToPDF', print_options)
            

            # 生成的 PDF 文件保存路径
            pdf_path = destination_folder + r'\output.pdf'

            # 保存为 PDF 文件
            with open(pdf_path, 'wb') as f:
                # f.write(bytes(pdf['data'], 'utf-8'))
                f.write(base64.b64decode(pdf['data']))

            # sleep(300)  # 暂停 5 分钟
            
        driver.quit()

        print(f"DESTINATION_FOLDER: {destination_folder}")
        print("finish")


    except Exception as e:
        print(f"err: {e}")
        sys.exit(1)
        
if __name__ == '__main__':
    main()
