
import base64
import pandas as pd
import json
import sys
import os
from time import sleep
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

CONFIG_PATH = r"C:\Users\beite.dai.bp\OneDrive - Coca-Cola Bottlers Japan\デスクトップ\py\automation_scheduler\config.json"
with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = json.load(f)

import_path = sys.argv[1]
kmcd = config["kmcd"]
userid = config["userid"]
userpw = config["userpw"]
destination_folder = config["destination_folder"]
subfolder_name = " WEB勤代修正"
destination_folder = os.path.join(destination_folder, subfolder_name)
os.makedirs(destination_folder, exist_ok=True)

# webdriver-manager 会自动下载浏览器对应版本webdriver
service = Service()
options = Options()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 10)

# 主功能模块
def WEB登録(url, userid, userpw):
    
    # 调用时传入 URL 和用户凭据
    driver.get(url)
    # 为啥最大化
    driver.maximize_window()
    # find and send
    login_id = driver.find_element(by="name", value="uid")
    login_pw = driver.find_element(by="name", value="pwd")
    login_id.send_keys(userid) 
    login_pw.send_keys(userpw)
    
    login_btn = driver.find_element(by="name", value="Login")
    login_btn.click()
    
    element = driver.find_element(by="css selector", value="span[title='勤怠システム/Attendance system']")
    element.click()

    #  分辨率and屏幕大小影响div变span
    # //*[@id="portal-menu-gadget"]/div/div[2]/div/div/div/div[5]/div/div[1]/div/div/div
    # //*[@id="portal-menu-gadget"]/div/div[2]/div/div/div/div[5]/div/div[1]/div/div/span
    # 元素加载慢
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@class='menu-list-item-title' and contains(text(), '主管部用メニュー')]")))
    element.click()
    # xpath css selector by.id 对比 精确定位？
    element = driver.find_element(by="css selector", value="span[title='勤務実績管理/Work record mgmt']")
    # element = driver.find_element(by="xpath", value="//span[@title='勤務実績管理/Work record mgmt']")
    element.click()

def 社員番号の登録(select_month, select_day,employee_number):

    # 定位到修改月 Select 选择
    select_element = driver.find_element(by="id", value="@PSTDDATEMONTH")
    select = Select(select_element)
    select.select_by_value(select_month)

    # 定位到修改日
    select_element = driver.find_element(by="id", value="@PSTDDATEDAY")
    select = Select(select_element)
    select.select_by_value(select_day)

    # input_box = driver.find_element(By.ID, "@PSTDDATEYEAR")
    # input_box.clear()  # 清空当前值
    # input_box.send_keys("2025")

    # 找到输入框并输入传入社员番号
    input_box = driver.find_element(
    by="xpath",
    value="/html/body/table[2]/tbody/tr/td/table/tbody/tr/td/form/div[2]/table/tbody/tr[2]/td[2]/div/div[2]/div/input")

    input_box.clear()

    input_box = driver.find_element(
    by="xpath",
    value="/html/body/table[2]/tbody/tr/td/table/tbody/tr/td/form/div[2]/table/tbody/tr[2]/td[2]/div/div[2]/div/input")

    input_box.send_keys(employee_number)
    sleep(2)
    # 按下回车键
    input_box.send_keys(Keys.RETURN)
    sleep(2)
    # 使用 XPath 或 ID 定位该按钮并点击
    button = driver.find_element(by="xpath", value="//input[@type='button' and @id='BTNLD']")
    button.click()

    sleep(2)
    # 使用 XPath 或 ID 定位该按钮并点击
    button = driver.find_element(by="xpath",
    value="/html/body/table[2]/tbody/tr/td/table/tbody/tr/td/form/div[3]/table[1]/tbody/tr[2]/td[3]/input")
    button.click()
    sleep(2)

def 勤怠の修正(formatted_date,kmcd):

    select_element = driver.find_element(By.NAME, "TGTDTRNG_SDT")

    # 初始化 Select 对象
    select = Select(select_element)

    # 选择指定日期（值为 2024_11_4）
    select.select_by_value(formatted_date)

    select_element = driver.find_element(By.NAME, "TGTDTRNG_EDT")

    # 初始化 Select 对象
    select = Select(select_element)

    # 选择指定日期
    select.select_by_value(formatted_date)

    driver.execute_script("document.getElementById('srw_fixed_footer_button_area').style.display = 'none';")
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.ID, "DLTDD")))


    # 点击按钮
    button = driver.find_element(By.ID, "DLTDD")
    button.click()

    sleep(2)
    button = driver.find_element(By.NAME, "btnExec0")
    button.click()

    sleep(2)
    link_element = driver.find_element(By.LINK_TEXT, "一覧画面へ（送信した勤怠期間）")
    link_element.click()
    sleep(2)

    #######################################
    KCDDT = "KCD" + formatted_date + "0S"
    print(KCDDT)
    select_element = Select(driver.find_element(By.NAME, KCDDT))
    # 根据 value 选择 "所定休日" 选项
    select_element.select_by_value(kmcd)

    # 等待触发操作完成
    sleep(2)                  
    #保存
    BTNDCDS2024_11_60 = "BTNDCDS" + formatted_date + "0"
    button = driver.find_element(By.NAME, BTNDCDS2024_11_60)
    button.click()
    sleep(2)
    #确认
    button = driver.find_element(By.NAME, "dSubmission0")
    button.click()

    sleep(3)

def 印刷PDF(destination_folder, pdf_file_name="output.pdf"):
    # 设置打印 PDF 的参数
    settings = {
        "paperWidth": 33.1,
        "paperHeight": 46.8,
        "marginTop": 0,
        "marginBottom": 0,
        "marginLeft": 0,
        "marginRight": 0,
        "scale": 0.75,
        "printBackground": True,
        "landscape": True
    }
    

    result = driver.execute_cdp_cmd("Page.printToPDF", settings)
    pdf_path = os.path.join(destination_folder, pdf_file_name)

    # 保存为 PDF 文件
    with open(pdf_path, "wb") as f:
        f.write(base64.b64decode(result['data']))

def 月次申請():

            button = wait.until(EC.presence_of_element_located((By.ID, "BTNSBMT0")))
            # 滚动页面，确保按钮在可见区域
            driver.execute_script("arguments[0].scrollIntoView();", button)
            sleep(0.5)  # 等待滚动完成
            # 用 JavaScript 直接点击，绕过点击拦截
            driver.execute_script("arguments[0].click();", button)
            sleep(1)  # 等待操作完成
            button = wait.until(EC.element_to_be_clickable((By.NAME, "btnExec0")))
            button.click()
            sleep(1)

            link_element = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "一覧画面へ（送信した勤怠期間）")))
            link_element.click()
            sleep(2)

            while True:
                try:
                    button = wait.until(EC.presence_of_element_located((By.ID, "BTNAPRV0")))
                    # 滚动页面，确保按钮在可见区域
                    driver.execute_script("arguments[0].scrollIntoView();", button)
                    sleep(0.5)  # 等待滚动完成
                    # 用 JavaScript 直接点击，绕过点击拦截
                    driver.execute_script("arguments[0].click();", button)
                    sleep(0.5)
                    button = wait.until(EC.element_to_be_clickable((By.NAME, "btnExec0")))
                    button.click()
                    sleep(1)
                    link_element = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "一覧画面へ（送信した勤怠期間）")))
                    link_element.click()
                    sleep(1)
                except TimeoutException:
                    print("按钮不存在，结束循环")
                    break  # 按钮找不到，跳出循环
                except Exception as e:
                    print(f"发生错误: {e}")
                    break  # 其他异常，跳出循环

            element = driver.find_element(By.ID, "BTNBCK2")
            driver.execute_script("arguments[0].click();", element)

            sleep(3)

sys.stdout.flush()

def main():
    try:
        sys.stdout.flush()
        # 主程序
        WEB登録('11', userid, userpw)#网址url，ID，PW
        print("Open website")
        df = pd.read_excel(import_path)
        df.loc[len(df)] = ['2025-01-19', 'endend']  # 确保和原来的列数一致

        df['日付'] = pd.to_datetime(df['日付']).dt.strftime('%Y-%m-%d')  # 只保留日期部分
        data = list(zip(df['日付'], df['社員番号']))

        previous_employee_number = None# 初始化变量

        # 假设 data 是所有数据列表，kmcd 在其他地方已定义
        for i, (date, employee_number) in enumerate(data):
            # 拆分日期字符串
            year, month, day = date.split("-")
            select_month = str(int(month))  # 去掉月份前导零
            search_day = str(int(day))        # 去掉日期前导零
            formatted_date = f"{year}_{int(month)}_{int(day)}"  # 格式化日期

            # 判断是否为员工编号的第一次出现
            if i == 0 or employee_number != data[i - 1][1]:
                # 首次出现：注册员工编号，并导出“修正前”PDF
                社員番号の登録(select_month, search_day, employee_number)
                print(destination_folder,f"Processing: {employee_number}")
                印刷PDF(f"修正前_{employee_number}_{formatted_date}.pdf")

            # 执行考勤修正（每行都需要执行）
            勤怠の修正(formatted_date, kmcd)

            # 判断是否为该员工编号的最后一次出现
            # 条件：当前行是最后一行，或下一行的员工编号与当前不同
            if i == len(data) - 1 or employee_number != data[i + 1][1]:
                印刷PDF(f"修正後_{employee_number}_{formatted_date}.pdf")
                # 如果需要在员工末次出现时还执行月次申請，可在此调用：
                月次申請()

            previous_employee_number = employee_number# 更新前一行的社员番号
        print("Finish")
        driver.quit()

    except Exception as e:
        print(f"err: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()


# web勤代修正手作业逻辑
# 库selenium 的方法(method) get find(name id xpath css selector)send click Select  库time
# service options 的配置
# 语法用chatgpt编程
# def函数(function) 模块化编程 
# 演示vsc的调试模式
# 虚拟环境*
