# encoding: utf-8

"""杂项测试"""

from openpyxl_extend import load_workbook_ex

def test_get_defined_names():
    """测试读取名称"""
    book = load_workbook_ex('test/test.xlsx')

    foo_range = book.defined_names['Hello']
    dests = foo_range.destinations

    for title, coord in dests:
        print title, coord

def test_read_table():
    """测试加载xlsx"""
    book = load_workbook_ex('test/test.xlsx')
    sheet = book.get_sheet_by_name('foo')
    table = sheet.get_table_by_name('Names')

    print table.name
    cell = table.get_cell(0, 0)
    print cell.value, cell.data_type
    print table.get_row_count(), table.get_col_count()

def main():
    """ 主入口 """
    test_read_table()

if __name__ == '__main__':
    main()
