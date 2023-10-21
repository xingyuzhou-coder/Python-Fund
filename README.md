# Python-Fund

## 这是一个可以自动读取EXCEl文件中的基金代码，实时更新最新净值的小工具

### 导入需要的模块
```python
import requests
from bs4 import BeautifulSoup
import re
import numpy as np
import pandas as pd
import openpyxl
```

### 读取excel中基金的代号 

usecols=[3], 选取工作表第4列 //
sheet_name="债券基金投资资产" 选取工作表

```python
def read_excel_code():
    bondfund_code=pd.read_excel(r'/Users/zhouxingyu/Library/Mobile Documents/com~apple~CloudDocs/Xingyu Zhou/[5] Financial/投资数据库/Excel文档/金融资产配置文档.xlsx',usecols=[3],dtype=str,sheet_name="债券基金投资资产")
    bondfund_code_list=bondfund_code.values.tolist()
    print (bondfund_code_list)
    return bondfund_code_list
```

# 打开目标表格，打开目标表单
```python
wb=openpyxl.load_workbook(r'/Users/zhouxingyu/Library/Mobile Documents/com~apple~CloudDocs/Xingyu Zhou/[5] Financial/投资数据库/Excel文档/金融资产配置文档.xlsx')
ws = wb['债券基金投资资产']
```

# 取出净值放入单元格
```python
for i in range(1,np.size(arr)+1):
    distance=arr[i-1]
    # 从第2行开始，插入第5列
    ws.cell(row = i+1, column = 5).value =distance
# 保存操作
wb.save(r'/Users/zhouxingyu/Library/Mobile Documents/com~apple~CloudDocs/Xingyu Zhou/[5] Financial/投资数据库/Excel文档/金融资产配置文档.xlsx')
```
