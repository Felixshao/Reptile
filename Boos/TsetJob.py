import requests, re, json, openpyxl, os, time
from bs4 import BeautifulSoup
from config.ReadInterfaceConfig import ReadInter

data = ReadInter().readExcel_config()[0]
params = eval(data['参数'])
headers = eval(data['请求头'])


def get_testJob_ing():
    """获取所有软件测试的招聘信息"""
    boos_url = 'https://www.zhipin.com'
    test_jobs = []
    for num in range(1, 15):
        params['page'] = str(num)
        test = requests.get(data['接口链接'], params=params, headers=headers)
        soup = BeautifulSoup(test.text, 'html.parser')
        for test in soup.find_all(name='div', attrs={'class': 'job-primary'}):
            job = {}
            exper = re.search('<p>(.+)<em class="vline"></em>(.*?)</p>', str(test))
            job['职位'] = test.find(name='span', attrs={'class': 'job-name'}).text
            job['公司'] = test.find_all(name='h3', attrs={'class': 'name'})[1].text
            job['经验'] = exper.groups()[0]
            job['学历'] = exper.groups()[1]
            job['地址'] = test.find(name='span', attrs={'class': 'job-area'}).text
            job['薪水'] = test.find(name='span', attrs={'class': 'red'}).text
            job['招聘链接'] = boos_url + test.find(name='div', attrs={'class': 'primary-box'})['href']
            test_jobs.append(job)

    write_excel('d:\\study\\python\\Reptile\\file\\BoosJob.xlsx', '软件测试2', test_jobs)


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
        sheet.cell(2, 1, '职位')
        sheet.cell(2, 2, '公司')
        sheet.cell(2, 3, '经验')
        sheet.cell(2, 4, '学历')
        sheet.cell(2, 5, '地址')
        sheet.cell(2, 6, '薪水')
        sheet.cell(2, 7, '招聘链接')
    row = 3

    for data in datas:
        col = 1
        for i in data.values():
            sheet.cell(row, col, i)
            col += 1
        row += 1
    job.save(tablepath)


if __name__ == '__main__':
    get_testJob_ing()
