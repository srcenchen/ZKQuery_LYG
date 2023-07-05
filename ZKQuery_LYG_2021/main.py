import requests

name = input("请输入姓名: ")
className = input("请输入班级,(例如1班，则输入 1): ")
for index in range(0, 60):
    if index < 10:
        number = '0{0}'.format(index)
    else:
        number = '{0}'.format(index)
    if int(className) < 10:
        className = '0{0}'.format(className)

    numberJDH = '225014{0}{1}'.format(className, number)
    response = requests.get('http://121.229.42.255:8888/grade/getGrade'
                            '?param={0}&password=0&captcha=372178'.format(numberJDH + "__" + name))
    if response.json()['code'] != 1002:
        data = response.json()['data']
        SW = data['SW']
        DL = data['DL']
        ZF = data['ZF']
        print('姓名：{0}\n建档号：{1}\n生物：{2}\n地理：{3}\n总分：{4}'
              .format(name, numberJDH, SW, DL, ZF))
        exit(0)
