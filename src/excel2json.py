# encoding: utf-8

"""
将excel文件中的table到导出成json对象
"""

import sys
import json
import os
from openpyxl_extend import load_workbook_ex


def _get_tables_to_parse(workbook, table_names):
    """ 根据table获得需要解析的table对象 """

    if table_names is None:
        return workbook.get_all_tables()
    else:
        all_tables = workbook.get_all_tables()
        tables = []
        for table_name in table_names:
            for table in all_tables:
                if table.name == table_name:
                    tables.append(table)
                    break
        return tables


def table2obj(table):
    """
    将table转换成python对象
    对于excel中的table
    列A | 列B
    行1A | 行1B
    行2A | 行2B
    将转换成
    [
        {"列A" : 行1A, "列B" : 行1B},
        {"列A" : 行2A, "列B" : 行2B}
    ]
    """

    row_count = table.get_row_count()
    col_count = table.get_col_count()

    # 获得列名
    col_names = []
    for col_id in range(col_count):
        col_names.append(table.get_cell(0, col_id).value)

    # 解析每一行
    result = []
    for row_id in xrange(1, row_count):
        row = {}
        for col_id in range(col_count):
            row[col_names[col_id]] = table.get_cell(row_id, col_id).value
        result.append(row)

    return result


def excel2json(filename, directory='.', table_names=None):
    """
    将filename对应的excel文件中的table导出成json文件
    如果table_names为None,则导出所有table,否则只导出table_names对应的table
    """

    workbook = load_workbook_ex(filename)
    tables = _get_tables_to_parse(workbook, table_names)
    for table in tables:
        obj = table2obj(table)
        json_str = json.dumps(obj, sort_keys=True,
                              indent=2, ensure_ascii=False)
        json_filename = os.path.splitext(os.path.basename(filename))[
            0] + '.' + table.name + '.json'
        fp_write = file(os.path.join(directory, json_filename), 'w')
        fp_write.write(json_str.encode('UTF-8'))
        fp_write.close()


USAGE = """
Usage:
    python excel2json excel_file_path out_dir [table_name1] [table_name2] ...

Example:
    python excel2json test.xlsx .
    python excel2json test.xlsx output Names Stages
"""


def usage():
    """ excel2json的命令行使用说明 """
    return USAGE


def cmd_main():
    """ 命令行主函数 """

    # 解析命令行参数
    if len(sys.argv) < 3:
        print usage()
        return

    # 执行操作
    excel_file = sys.argv[1]
    out_dir = sys.argv[2]

    if len(sys.argv) == 3:
        excel2json(excel_file, out_dir)
    else:
        table_names = []
        for i in xrange(3, len(sys.argv)):
            table_names.append(sys.argv[i])
        excel2json(excel_file, out_dir, table_names)


if __name__ == '__main__':
    cmd_main()
