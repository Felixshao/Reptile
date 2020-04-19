import requests, re, openpyxl, os, time, warnings
from bs4 import BeautifulSoup
from config.ReadInterfaceConfig import ReadInter

warnings.filterwarnings('ignore')        # 忽略warnings警告
data = ReadInter().readExcel_config()[1]    # 获取csdn搜索appscan接口信息
params = eval(data['参数'])
headers = eval(data['请求头'])


def get_csdn_appscan():
    """获取scdn中appscan学习教程"""
    appscan_studys = []
    for i in range(1, 11):
        params['p'] = str(i)
        appscan = requests.get(data['接口链接'], params=params, headers=headers, verify=False)
        soup = BeautifulSoup(appscan.text, 'html.parser')
        for i in soup.find_all(name='dl', attrs={'class': 'search-list J_search'}):
            appscan = {}
            study = re.search('<a href=(.+) target="_blank">(.*?)</a>', str(i))
            author_time = i.find(name='dd', attrs={'class': 'author-time'})
            appscan['教程名称'] = study.groups()[1].replace('<em>', '').replace('</em>', '').replace(' ', '')
            appscan['教程链接'] = study.groups()[0].replace('"', '')
            appscan['作者'] = author_time.a.string
            appscan['浏览量/大小'] = i.find(name='span', attrs={'class': 'mr16'}).text
            appscan['时间'] = i.find(name='span', attrs={'class': 'date'}).text
            appscan_studys.append(appscan)

    write_excel('d:\\study\\python\\Reptile\\file\\appscan_csdn.xlsx', 'appscan_csdn2', appscan_studys)


def write_excel(tablepath, sheetname, datas):
    """写入excel表"""
    if os.path.exists(tablepath):
        job = openpyxl.load_workbook(tablepath)
    else:
        job = openpyxl.Workbook(tablepath)
    if sheetname in job.sheetnames:
        sheet = job[sheetname]
    else:
        job.create_sheet(title=sheetname)
        job.save(tablepath)
        time.sleep(3)
        job = openpyxl.load_workbook(tablepath)
        sheet = job[sheetname]
        sheet.cell(1, 1, sheetname)
        sheet.cell(2, 1, '教程名称')
        sheet.cell(2, 2, '教程链接')
        sheet.cell(2, 3, '作者')
        sheet.cell(2, 4, '浏览量')
        sheet.cell(2, 5, '时间')
    row = 3
    for data in datas:
        col = 1
        for i in data.values():
            sheet.cell(row, col, str(i))
            col += 1
        row += 1
    job.save(tablepath)


if __name__ == '__main__':
    get_csdn_appscan()