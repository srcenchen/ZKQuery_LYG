import requests
import ddddocr
import pandas, json

# 保存
import xlrd
from xlutils.copy import copy


def save_to_excel(listz, path):
    value = list(listz.values())
    index = 1  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(listz)):
            new_worksheet.write(i + rows_old, j, value[j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    # print("xls/xlsx格式表格【追加】写入数据成功！")


# 查询
def search(jdh, password, final_bmh, class_name, name):
    url = "http://121.229.42.255:8996/grade/getCaptcha"
    result = requests.get(url)
    cookies = result.cookies
    with open('captcha.png', 'wb') as file:
        file.write(result.content)

    ocr = ddddocr.DdddOcr(show_ad=False)
    with open('captcha.png', 'rb') as f:
        img_bytes = f.read()
    captcha = ocr.classification(img_bytes)
    # 稍微验证下验证码
    if len(captcha) > 5:
        captcha = captcha[1:]
    # print(captcha)
    res = json.loads(requests.get(
        "http://121.229.42.255:8996/grade/getGrade?param=" + str(jdh) + "_" + str(final_bmh) + "_&password=" + "0" + str(
            password) + "&captcha=" + captcha,
        cookies=cookies).text)
    print("查询" + str(jdh) + "_" + str(final_bmh) + str(res))
    code = res["code"]
    data = res["data"]
    if code == 1001:
        # 意味着验证码错误，我们要再次尝试
        print("验证码错误，再次尝试" + str(jdh) + "_" + str(final_bmh))
        search(jdh, password, final_bmh, class_name, name)
    if code == 0:
        # 意味着查询成功了 写入成绩
        print(res["data"])  # 这里记得改为 res
        result_dict = res["data"]  # 这里记得改为 res
        result_dict["BJ"] = class_name
        result_dict["XM"] = name
        save_to_excel(result_dict, "./结果/成绩.xls")
    if code != 1001 and data is None:
        # 异常，记录在案
        save_to_excel({'姓名': name, '建档号': jdh, '报名号': final_bmh, '密码': password, 'msg': res["msg"]}, "./结果/异常名单.xls")


# Excel 读取
def excel_read():
    result = pandas.read_excel("./final.xls", sheet_name='Sheet1')
    list_result = result.values.tolist()
    for list_item in list_result:
        jdh = list_item[1]
        password = list_item[3]
        final_bmh = list_item[2]
        class_name = list_item[4]
        name = list_item[0]
        search(jdh, password, final_bmh, class_name, name)
        # print(jdh, password, final_bmh, class_name, name)


if __name__ == '__main__':
    excel_read()
