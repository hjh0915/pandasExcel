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

    def get_total_03(self, late_time='否'):
        """计算缓装数量"""
        t = self.df[self.df['是否缓装'] == late_time].groupby(['区县'], as_index=False).size()
        results = []
        for i, r in t.iterrows():
            d = dict()
            d['cc'] = r[0]
            d['late_time'] = late_time
            d['cnt'] = r[1]
            results.append(d) 
        return results

class EFrame(object):
    def __init__(self):
        pass

    def gen_data_01(self, result, item_type):
        """构建区县、类型dataframe"""
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
        """构建区县、类型、每个类型的超时单dataframe"""
        res = [x for x in result if x['item_type'] == item_type]
        cc = [x['cc'] for x in res]
        cnt = [x['cnt'] for x in res]
        data = {
            '区县': cc,
            '超时单': cnt
        }
        return pd.DataFrame(data)

    def gen_data_03(self, result):
        """构建区县、缓装dataframe"""
        cc = [x['cc'] for x in result]
        cnt = [x['cnt'] for x in result]
        data = {
            '区县': cc,
            '缓装工单': cnt
        }
        return pd.DataFrame(data)

def merge_data(df1, df2):
    """合并两个DataFrame"""
    df = pd.merge(df1, df2, on='区县')
    return df

def create_by_items(item_type1, item_type2, item_type3):
    """根据不同的item_type分别构建dataframe"""
    # 新装
    result1 = ed.get_total_01()
    data1 = ef.gen_data_01(result1, item_type1)
    
    result1 = ed.get_total_02(item_type1, over_time='是')
    new_result1 = fill_data(result1)
    data2 = ef.gen_data_02(new_result1, item_type1)

    df1 = merge_data(data1, data2)

    # 移机
    result2 = ed.get_total_01()
    data3 = ef.gen_data_01(result2, item_type2)
    
    result2 = ed.get_total_02(item_type2, over_time='是')
    new_result2 = fill_data(result2)
    data4 = ef.gen_data_02(new_result2, item_type2)

    df2 = merge_data(data3, data4)

    # 故障单
    result3 = ed.get_total_01()
    data5 = ef.gen_data_01(result3, item_type3)
    
    result3 = ed.get_total_02(item_type3, over_time='是')
    new_result3 = fill_data(result3)
    data6 = ef.gen_data_02(new_result3, item_type3)

    df3 = merge_data(data5, data6)

    all_df1 = pd.merge(df1, df2, on='区县')
    all_df = pd.merge(all_df1, df3, on='区县')

    return all_df

def write_excel(filename, df):
    """生成一张Excel表"""
    writer = pd.ExcelWriter(filename)
    df.to_excel(writer, 'Sheet1')
    writer.save()

if __name__ == '__main__':
    file = 'initdata/details.xlsx'
    ed = EData(file)
    ef = EFrame()

    item_type1 = '新装'
    item_type2 = '移机'
    item_type3 = '故障单'
    items_df = create_by_items(item_type1, item_type2, item_type3)

    # 添加缓装工单数据
    result = ed.get_total_03(late_time='是')
    late_time_df = ef.gen_data_03(result)

    all_df = pd.merge(items_df, late_time_df, on='区县')

    # 在汇总表中添加计算三种类型的数据总计列
    all_df['总计'] = all_df[['新装', '移机', '故障单']].apply(lambda x: x['新装']+x['移机']+x['故障单'], axis=1)

    # 在汇总表中添加计算的超时单数据总计列
    all_df['超时单总计'] = all_df[['超时单_x', '超时单_y', '超时单']].apply(lambda x: x['超时单_x']+x['超时单_y']+x['超时单'], axis=1)

    # 在汇总表中添加计算的超时单对于总计的占比列
    all_df['超时单占比总计'] = all_df[['超时单总计', '总计']].apply(lambda x: x['超时单总计']/x['总计'], axis=1)
    all_df['超时单占比总计'] = all_df['超时单占比总计'].apply(lambda x:format(x, '.2%'))

    write_excel('data/total.xlsx', all_df)
