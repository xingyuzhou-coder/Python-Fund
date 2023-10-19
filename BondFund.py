# 导入需要的模块
import requests
from bs4 import BeautifulSoup
import re
import numpy as np
import pandas as pd
import openpyxl

# 1. 抓取网页
def get_url(url, params=None, proxies=None):
    rsp = requests.get(url, params=params, proxies=proxies)
    rsp.raise_for_status()
    return rsp.text

# 2. 从网页抓取数据
def get_fund_data(code,per=10,sdate='',edate='',proxies=None):
    url = 'http://fund.eastmoney.com/f10/F10DataApi.aspx'
    params = {'type': 'lsjz', 'code': code, 'page':1,'per': per, 'sdate': sdate, 'edate': edate}
    html = get_url(url, params, proxies)
    print(html)
    soup = BeautifulSoup(html, 'html.parser')

    # 获取总页数
    pattern=re.compile(r'pages:(.*),')
    result=re.search(pattern,html).group(1)
    pages=int(result)

    # 获取表头
    heads = []
    for head in soup.findAll("th"):
        heads.append(head.contents[0])

    # 数据存取列表
    records = []

    # 从第1页开始抓取所有页面数据
    page=1
    while page<=pages:
        params = {'type': 'lsjz', 'code': code, 'page':page,'per': per, 'sdate': sdate, 'edate': edate}
        html = get_url(url, params, proxies)
        soup = BeautifulSoup(html, 'html.parser')

        # 获取数据
        for row in soup.findAll("tbody")[0].findAll("tr"):
            row_records = []
            for record in row.findAll('td'):
                val = record.contents

                # 处理空值
                if val == []:
                    row_records.append(np.nan)
                else:
                    row_records.append(val[0])

            # 记录数据
            records.append(row_records)

        # 下一页
        page=page+1

    # 数据整理到dataframe
    np_records = np.array(records)
    data= pd.DataFrame()
    for col,col_name in enumerate(heads):
        data[col_name] = np_records[:,col]

    return data

# 读取excel中基金的代号
def read_excel_code():
    bondfund_code=pd.read_excel(r'/Users/zhouxingyu/Library/Mobile Documents/com~apple~CloudDocs/Xingyu Zhou/[5] Financial/投资数据库/Excel文档/金融资产配置文档.xlsx',usecols=[3],dtype=str,sheet_name="债券基金投资资产")
    bondfund_code_list=bondfund_code.values.tolist()
    print (bondfund_code_list)
    return bondfund_code_list

# 主程序
if __name__ == "__main__":
    # 读取需要写的数据代码
    alldata = read_excel_code()
    arr = []
    # 获取净值
    for code in alldata:
        print(code)
        data = get_fund_data(code, per=49,sdate='2023-1-1')  # sdate=开始日期
        # 修改数据类型
        data['净值日期']=pd.to_datetime(data['净值日期'],format='%Y-%m-%d')
        data['单位净值']= data['单位净值'].astype(float)
        data['累计净值']=data['累计净值'].astype(float)
        data['日增长率']=data['日增长率'].str.strip('%').astype(float)
        # 按照日期升序排序并重建索引
        data=data.sort_values(by='净值日期',axis=0,ascending=False).reset_index(drop=True)
        # 获取净值日期、单位净值、累计净值、日增长率等数据
        net_value_date = data['净值日期']
        net_asset_value = data['单位净值']
        accumulative_net_value=data['累计净值']
        daily_growth_rate = data['日增长率']

        arr = np.append(arr, net_asset_value[0], axis=None)
        print(arr)

'''
将净值按列插入目标表格的某列
'''
# 打开目标表格，打开目标表单
wb=openpyxl.load_workbook(r'/Users/zhouxingyu/Library/Mobile Documents/com~apple~CloudDocs/Xingyu Zhou/[5] Financial/投资数据库/Excel文档/金融资产配置文档.xlsx')
ws = wb['债券基金投资资产']

# 取出净值放入单元格
for i in range(1,np.size(arr)+1):
    distance=arr[i-1]
    # 从第2行开始，插入第5列
    ws.cell(row = i+1, column = 5).value =distance
# 保存操作
wb.save(r'/Users/zhouxingyu/Library/Mobile Documents/com~apple~CloudDocs/Xingyu Zhou/[5] Financial/投资数据库/Excel文档/金融资产配置文档.xlsx')

print('\n\n')
print("撒娇可爱聪明卖萌小齐齐计算完毕，快来夸我吧!")