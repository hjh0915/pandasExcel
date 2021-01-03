pandas+Excel表格
================
从Excel中获取数据
---------------
```py
def __init__(self, filename):
    xl_file = pd.ExcelFile(filename)
    self.df = xl_file.parse('Sheet1')
```

再写入并生成Excel表格
-------------------
```py
def write_excel(filename, df):
    """生成一张Excel表"""
    writer = pd.ExcelWriter(filename)
    df.to_excel(writer, 'Sheet1')
    writer.save()
```

不取字段行
========
```py
df.groupby(['区县','工单类型'], as_index=False).size()
```
其中使用到了as_index=False方法

将pandas转换成python的数据结构
===========================
```py
results = []
for i, r in t.iterrows():
    d = dict()
    d['cc'] = r[0]
    d['item_type'] = r[1]
    d['cnt'] = r[2]
    results.append(d) 
return results
```

将python转换成pandas结构
======================
转换成dataframe形式
```py
res = [x for x in result if x['item_type'] == item_type]
cc = [x['cc'] for x in res]
item = [x['item_type'] for x in res]
cnt = [x['cnt'] for x in res]
data = {
    '区县': cc,
    item_type: cnt
}
return pd.DataFrame(data)
```