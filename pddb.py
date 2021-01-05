import pandas as pd 

all_cc = [
    '城区分局', '丰城市', '奉新县', '高安市', '靖安县', '上高县', 
    '铜鼓县', '万载县', '宜丰县', '袁州分局', '樟树市'
]

WORK_ORDER = ['新装', '移机', '故障单']

def fill_data(result):
    """补充表单中为0的数据"""
    s = []
    cc = [r['cc'] for r in result]
    item_type = result[0]['item_type']
    over_time = result[0]['over_time']
    for x in all_cc:
        if x not in cc:
            d = dict()
            d['cc'] = x
            d['item_type'] = item_type
            d['over_time'] = over_time
            d['cnt'] = 0 
            s.append(d)
    return result + s

class EData(object):
    def __init__(self, filename):
        xl_file = pd.ExcelFile(filename)
        self.df = xl_file.parse('Sheet1')

    def get_total_01(self):
        """区县汇总"""
        t = self.df.groupby(['区县','工单类型'], as_index=False).size()
        results = []
        for i, r in t.iterrows():
            d = dict()
            d['cc'] = r[0]
            d['item_type'] = r[1]
            d['cnt'] = r[2]
            results.append(d) 
        return results

    def get_total_02(self, item_type, over_time='否'):
        """按工单类型过滤汇总"""
        t = self.df[(self.df['工单类型'] == item_type) & (self.df['是否超时'] == over_time)].groupby(['区县', '工单类型'], as_index=False).size()
        results = []
        for i, r in t.iterrows():
            d = dict()
            d['cc'] = r[0]
            d['item_type'] = r[1]
            d['over_time'] = over_time
            d['cnt'] = r[2]
            results.append(d) 
        return results

class EFrame(object):
    def __init__(self):
        pass

    def gen_data_01(self, result, item_type):
        res = [x for x in result if x['item_type'] == item_type]
        cc = [x['cc'] for x in res]
        item = [x['item_type'] for x in res]
        cnt = [x['cnt'] for x in res]
        data = {
            '区县': cc,
            item_type: cnt
        }
        return pd.DataFrame(data)


    def gen_data_02(self, result, item_type):
        res = [x for x in result if x['item_type'] == item_type]
        cc = [x['cc'] for x in res]
        item = [x['item_type'] for x in res]
        over_time = [x['over_time'] for x in res]
        cnt = [x['cnt'] for x in res]
        data = {
            '区县': cc,
            '超时单': cnt
        }
        return pd.DataFrame(data)

def merge_data(df1, df2):
    """合并两个DataFrame"""
    df = pd.merge(df1, df2, on='区县')
    return df

def write_excel(filename, df):
    """生成一张Excel表"""
    writer = pd.ExcelWriter(filename)
    df.to_excel(writer, 'Sheet1')
    writer.save()

if __name__ == '__main__':
    file = 'initdata/details.xlsx'
    ed = EData(file)
    ef = EFrame()

    item_type = '新装'
    result = ed.get_total_01()
    data1 = ef.gen_data_01(result, item_type)
    
    result = ed.get_total_02(item_type, over_time='是')
    new_result = fill_data(result)
    data2 = ef.gen_data_02(new_result, item_type)

    df = merge_data(data1, data2)

    write_excel('data/total.xlsx', df)
