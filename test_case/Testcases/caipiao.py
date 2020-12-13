#!/usr/bin/python3
# -*- coding:utf-8 -*-
# @Author:tan
# @Time:2019/11/21 14:08
# @Email:1164355091@qq.com
import unittest
import requests
from Unittest_Caipiao.Common import log,excel_reading
from ddt import ddt,data,unpack
# test_data=[{"key":"d86a380963d6cc16ad80445d70997156",
#               "lottery_id":"ssq"},
#            {"key": "d86a380963d6cc16ad80445d70997156",
#             "lottery_id": "dlt"},
#            {"key":"d86a380963d6cc16ad80445d70997156",
#               "lottery_id":"fcsd"}]
excel=excel_reading.Excel(r"E:\python代码\Unittest_Caipiao\Testdatas\test.xlsx", "Sheet1")
test_data=excel.row_col_appointed_tuple([2,3,4,5], [1,5,6,7])
max_col=excel.max_column()
print(test_data)

@ddt
class CaiPiao(unittest.TestCase):
    def setUp(self):
        self.log=log.log().get_log()
    @data(*test_data)
    @unpack
    def test_caipiao(self,hang,url,request_data,expected):
        #测试查询彩票彩种
        #操作步骤，实际结果，预期结果
        self.log.info("开始测试彩票彩种")
        r=requests.post(url,request_data)
        self.log.info("获取到了响应")
        rel=r.json()["reason"]
        excel.write_data(int(hang)+1,max_col-1,rel)
        try:
            self.assertEqual(rel,expected)
            result="通过"
            excel.write_data(int(hang)+1, max_col, result)
            self.log.info("获取彩票彩种信息成功")
            self.log.info("用例执行成功")
        except:
            result="未通过"
            excel.write_data(int(hang)+1, max_col, result)
            self.log.error("用例执行失败")