import requests, json, os, openpyxl, time, sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from config.ProjectPath import get_project_path
from config.ReadInterfaceConfig import ReadInter

stock_east_data = ReadInter().readExcel_config()[2]
path = get_project_path()
stock_east_url = stock_east_data['接口链接']
headers = eval(stock_east_data['请求头'])
params = eval(stock_east_data['参数'])
filepath = os.path.join(path, 'file', 'stock_east.xlsx')
stocks_sheetname = '东方财富股票统计'
stock_count_sheetname = '近两天连涨股票统计'


def get_stock_east():
    """东方财富网获取信息"""
    stock_lists = []
    for i in range(1, 6):
        params['pn'] = i
        stock = requests.get(stock_east_url, params=params, headers=headers)
        stock_datas = \
            eval(stock.text.replace('jQuery18306380779290944183_1587281762105', '').replace(';', ''))['data']['diff']
        for data in stock_datas:
            stock_dict = {}
            stock_dict['名称'] = data['f14']
            stock_dict['代码'] = data['f12']
            stock_dict['最新'] = '最新:' + str(data['f2']) + ', '
            stock_dict['跌涨幅'] = '跌涨幅:' + str(data['f3']) + '%, '
            stock_dict['今日主力净流入'] = '今日主力净流入:' + str(data['f62']) + ', '
            stock_dict['今日超大单净流入'] = '今日超大单净流入:' + str(data['f66'])
            stock_dict['时间'] = str(time.strftime('%Y%m%d', time.localtime()))
            stock_lists.append(stock_dict)
    write_excel(filepath, stocks_sheetname, stock_lists)


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
        sheet.cell(2, 1, '名称')
        sheet.cell(2, 2, '代码')
        sheet.cell(2, 3, datas[0]['时间'])

    max_row = sheet.max_row
    max_col = sheet.max_column
    if sheet[3][0].value is None:
            row = 3
            for data in datas:
                col = 1
                sheet.cell(row, col, data['名称'])
                sheet.cell(row, col + 1, data['代码'])
                sheet.cell(row, col + 2, data['最新'] + data['跌涨幅'] + data['今日主力净流入'] + data['今日超大单净流入'])
                row += 1
    else:
        if str(sheet[2][max_col - 1].value) != datas[0]['时间']:
            max_col += 1
            sheet.cell(2, max_col, datas[0]['时间'])
        old_stock_data = get_excel_stockCode(tablepath, sheetname)
        for data in datas:
            if data['代码'] in old_stock_data:
                row = old_stock_data.index(data['代码']) + 3
                sheet.cell(row, max_col, data['最新'] + data['跌涨幅'] + data['今日主力净流入'] + data['今日超大单净流入'])
            else:
                col = 1
                max_row += 1
                sheet.cell(max_row, col, data['名称'])
                sheet.cell(max_row, col + 1, data['代码'])
                sheet.cell(max_row, max_col, data['最新'] + data['跌涨幅'] + data['今日主力净流入'] + data['今日超大单净流入'])
    job.save(tablepath)


def get_excel_stockCode(filepath, sheetname):
    """获取excel表中以保存的股票代码，存入list"""
    job = openpyxl.load_workbook(filepath)
    sheet = job[sheetname]
    data_lists = []
    for row in range(3, sheet.max_row+1):
        data_lists.append(sheet[row][1].value)
    return data_lists


def stat_nearly_three():
    """统计近两天连涨股票"""
    table = openpyxl.load_workbook(filepath)
    sheet = table[stocks_sheetname]
    max_row = sheet.max_row
    max_col = sheet.max_column
    stocks = []
    for row in range(3, max_row+1):
        count = {}
        if sheet[row][max_col-1].value and sheet[row][max_col-2].value:
            if float(sheet[row][max_col-1].value.split(',')[0].split(':')[1]) <= 20.0:
                count['名称'] = sheet[row][0].value
                count['代码'] = sheet[row][1].value
                stocks.append(count)

    # 写入统计表格
    """写入excel表"""
    if os.path.exists(filepath):
        job = openpyxl.load_workbook(filepath)
    else:
        job = openpyxl.Workbook(filepath)
    if stock_count_sheetname in job.sheetnames:
        sheet = job[stock_count_sheetname]
    else:
        job.create_sheet(title=stock_count_sheetname)
        job.save(filepath)
        time.sleep(3)
        job = openpyxl.load_workbook(filepath)
        sheet = job[stock_count_sheetname]
        sheet.cell(1, 1, stock_count_sheetname)
        sheet.cell(2, 1, '名称')
        sheet.cell(2, 2, '代码')
    max_row = sheet.max_row
    if sheet[3][0].value is not None:
        for row in range(3, max_row+1):
            sheet.cell(row, 1, '')
            sheet.cell(row, 2, '')
    row = 3
    for data in stocks:
        sheet.cell(row, 1, data['名称'])
        sheet.cell(row, 2, data['代码'])
        row += 1
    job.save(filepath)


if __name__ == '__main__':
    get_stock_east()
    time.sleep(5)
    stat_nearly_three()

