#!/usr/bin/python3
# -*- coding:utf-8 -*-
# @Author:tan
# @Time:2019/11/21 15:40
# @Email:1164355091@qq.com
import os
import time
from Unittest_Caipiao.Output import getcwd
from Unittest_Caipiao.Common.HTMLTestRunnerNew import HTMLTestRunner
from Unittest_Caipiao.Testcases import caipiao
import  unittest

suit=unittest.TestSuite()
loader=unittest.TestLoader()
suit.addTest(loader.loadTestsFromTestCase(caipiao.CaiPiao))
run=unittest.TextTestRunner()
# 获取本地时间，转换为设置的格式
rq = time.strftime('%Y%m%d%H%M', time.localtime(time.time()))
# 设置所有测试报告的存放路径
path = getcwd.get_cwd()
# 通过getcwd.py文件的绝对路径来拼接报告存放路径
all_path = os.path.join(path, 'Reports/')

# 设置报告文件名
all_name = all_path + rq + '.html'

with open(all_name,'wb' ) as f:
    test=HTMLTestRunner(
                        stream=f,
                        verbosity=2,                #详细程度
                        title='测试报告',
                        description="第一份测试报告",
                        tester="tan")
    test.run(suit)