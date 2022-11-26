#引入selenium库中的 webdriver 模块
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from docx import Document
from docx.enum.section import WD_ORIENTATION

#DOS输入个人信息
a=input("请输入学号: ")
b=input("请输入学生姓名: ")
c=input("请输入导师：")

#当前版本Chrome忽略无用日志
options=webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])

#打开谷歌浏览器
driver = webdriver.Chrome(chrome_options=options,executable_path="D:\Selenium\chromedriver")
#打开登录页
driver.get('https://628251.yichafen.com/public/queryscore/sqcode/MszcQnwmNDI2OXw0YWU5NTYzZTIzNjIxYmRlZTk2MjkwMDU1NzA3ZTdmM3w2MjgyNTEO0O0O.html')

#在固定位置输入个人信息，并点击登录
driver.find_element(By.XPATH,r'//*[@id="queryForm"]/table/tbody/tr[1]/td[2]/input').send_keys(a)
driver.find_element(By.XPATH,r'//*[@id="queryForm"]/table/tbody/tr[2]/td[2]/input').send_keys(b)
driver.find_element(By.XPATH,r'//*[@id="queryForm"]/table/tbody/tr[3]/td[2]/input').send_keys(c)
driver.find_element(By.XPATH,r'//*[@id="yiDunSubmitBtn"]').click()
sleep(2)

document = Document()
#设置纸张方向为横向
section = document.sections[0]
new_width, new_height = section.page_height, section.page_width
section.orientation = WD_ORIENTATION.LANDSCAPE
section.page_width = new_width 
section.page_height = new_height

#生成一个表格
table = document.add_table(rows=10, cols=6)

#设置表头
uptitle = ["上课时间","星期一","星期二","星期三","星期四","星期五"]
for x in range(0,6):
    cell = table.cell(0,x)
    cell.text = uptitle[x]
lefttitle = ["上课时间","第1节\n08:00-08:45","第2节\n08:55-09:40","第3节\n09:50-10:35","第4节\n10:45-11:30","第5节\n11:40-12:25","第7节\n13:30-14:15","第8节\n14:25-15:10","第9节\n15:30-16:15","第10节\n4:40-5:15"]
for x in range(0,10):
    cell = table.cell(x,0)
    cell.text = lefttitle[x]

#从网页获取元素并填入表格
havedone = [1,1]
for x in range(4,45):
    if(x==20):
        continue
    xpath = '//*[@id="result_content"]/div[2]/table/tbody/tr[2]/td['+str(x)+']'
    now = driver.find_element(By.XPATH,xpath).text
    cell = table.cell(havedone[0],havedone[1])
    cell.text = now
    if(havedone[0]==8):
        havedone[1]+=1
        havedone[0]=1
        continue
    havedone[0]+=1
Tuesday = driver.find_element(By.XPATH,r'//*[@id="result_content"]/div[2]/table/tbody/tr[2]/td[20]').text
Tuesdaycell = table.cell(9,2)
Tuesdaycell.text = Tuesday

#保存文件到D盘
document.save(r"D:\你的课程表.docx")