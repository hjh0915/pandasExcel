import sys
sys.path.append('.')

import pddb
import unittest

class TestPdata(unittest.TestCase):
    """Tester for the function patients_with_missing_values in
    treatment_functions.
    """

    def setUp(self):
        file = 'initdata/details.xlsx'
        self.db = pddb.EData(file)
        self.pd = pddb.EFrame()

    def test_total_01(self):
        """测试区县汇总"""
        x = self.db.get_total_01()
        print(x)

    def test_total_02(self):
        """测试按照工单类型过滤汇总"""
        item_type = "新装"
        x = self.db.get_total_02(item_type, over_time='是')
        print(x)

        item_type = "移机"
        x = self.db.get_total_02(item_type, over_time='是')
        print(x)

        item_type = "故障单"
        x = self.db.get_total_02(item_type, over_time='是')
        print(x)

    def test_gen_data_01(self):
        """测试生成第一部分的data pandas"""
        result = self.db.get_total_01()
        item_type = '新装'
        data = self.pd.gen_data_01(result, item_type)
        print(data)

    def test_gen_data_02(self):
        """测试生成第二部分的data pandas"""
        item_type = "新装"
        result = self.db.get_total_02(item_type, over_time='是')
        new_result = pddb.fill_data(result)
        data = self.pd.gen_data_02(new_result, item_type)
        print(data)

    def test_merge_data(self):
        """测试生成一个合并的df"""
        item_type = '新装'
        result = self.db.get_total_01()
        data1 = self.pd.gen_data_01(result, item_type)
        
        result = self.db.get_total_02(item_type, over_time='是')
        new_result = pddb.fill_data(result)
        data2 = self.pd.gen_data_02(new_result, item_type)

        df = pddb.merge_data(data1, data2)
        print(df)

if __name__ == '__main__':
    unittest.main(exit=False)