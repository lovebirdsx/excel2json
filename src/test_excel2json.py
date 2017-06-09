# encoding: utf-8

"""
测试excel2json
"""

import os
import shutil
import unittest

from excel2json import table2obj, excel2json
from openpyxl_extend import load_workbook_ex


def clean_and_make_dir(dirname):
    """ 移除dirname对应的目录 """
    if os.path.exists(dirname):
        shutil.rmtree(dirname)
    os.mkdir(dirname)


class TestExcel2Json(unittest.TestCase):
    """
    测试excel2json
    """

    OUT_DIR1 = 'output1'
    OUT_DIR2 = 'output2'

    def __init__(self, *args, **kwargs):
        """ 测试的相关初始化 """

        super(TestExcel2Json, self).__init__(*args, **kwargs)

        clean_and_make_dir(TestExcel2Json.OUT_DIR1)
        clean_and_make_dir(TestExcel2Json.OUT_DIR2)

    def test_table2obj(self):
        """ 测试解析table """

        workbook = load_workbook_ex('test/test.xlsx')
        table = workbook.get_sheet_by_name('foo').get_table_by_name('Names')
        obj = table2obj(table)

        row0 = obj[0]
        self.assertEqual(row0['Name'], 'Hello')
        self.assertEqual(row0['Level'], 1)
        self.assertTrue(row0['Bool'])

    def test_excel2json_with_no_table(self):
        """ 测试将excel写入json文件, 不指定table """

        excel2json('test/test.xlsx', TestExcel2Json.OUT_DIR1)
        # test.xlsx中总共有3个Table,分别为Names, Levels, Stages
        self.assertTrue(os.path.exists(os.path.join(
            TestExcel2Json.OUT_DIR1, 'test.Names.json')))
        self.assertTrue(os.path.exists(os.path.join(
            TestExcel2Json.OUT_DIR1, 'test.Levels.json')))
        self.assertTrue(os.path.exists(os.path.join(
            TestExcel2Json.OUT_DIR1, 'test.Stages.json')))

    def test_excel2json_with_tables(self):
        """ 测试将excel写入json文件, 指定table """

        excel2json('test/test.xlsx', TestExcel2Json.OUT_DIR2, ['Names'])
        self.assertTrue(os.path.exists(os.path.join(
            TestExcel2Json.OUT_DIR2, 'test.Names.json')))
        self.assertFalse(os.path.exists(os.path.join(
            TestExcel2Json.OUT_DIR2, 'test.Levels.json')))
        self.assertFalse(os.path.exists(os.path.join(
            TestExcel2Json.OUT_DIR2, 'test.Stages.json')))


if __name__ == '__main__':
    unittest.main()
