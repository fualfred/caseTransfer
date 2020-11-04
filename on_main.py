#!/usr/bin/python
# -*- coding: utf-8 -*-
from utils import Utils as ut
import time

# level_dict = {
#     'P1': '1',
#     'P2': '2',
#     'P3': '3',
#     'P4': '4',
# }


def create_test_case(doc, test_msg):
    # print(test_msg)
    # print(test_msg)
    testcase = ut.create_element(doc, "testcase", None, test_msg[0])
    summary = ut.create_section(doc, "summary", test_msg[1].replace("\n", "<br>"))
    execution_type = ut.create_section(doc, "execution_type", str(1))
    importance = ut.create_section(doc, "importance",test_msg[2])
    preconditions = ut.create_section(doc, "preconditions", test_msg[3].replace("\n", "<br>"))
    steps = ut.create_element(doc, "steps", None)
    step = ut.create_element(doc, "step", None)
    ut.append_child_elem(steps, step)
    step_number = ut.create_section(doc, "step_number", str(1))
    actions = ut.create_section(doc, "actions", test_msg[4].replace("\n", "<br>"))
    expectedresults = ut.create_section(doc, "expectedresults", test_msg[5].replace("\n", "<br>"))
    ut.append_child_elem(step, step_number)
    ut.append_child_elem(step, actions)
    ut.append_child_elem(step, expectedresults)
    ut.append_child_elem(testcase, summary)
    ut.append_child_elem(testcase, execution_type)
    ut.append_child_elem(testcase, importance)
    ut.append_child_elem(testcase, preconditions)
    ut.append_child_elem(testcase, steps)
    return testcase


def add_test(doc, second_suit_value, test_msg, init_second_suits, init_suit):
    if second_suit_value is None:
        test_case = create_test_case(doc, test_msg)
        # append_suit = ut.get_element_by_attr_name(doc, "testsuite", first_suit_value)
        ut.append_child_elem(init_suit, test_case)
        return
    if second_suit_value not in init_second_suits:
        init_second_suits.append(second_suit_value)
        second_suit = ut.create_element(doc, "testsuite", None, second_suit_value)
        # append_suit = ut.get_element_by_attr_name(doc, "testsuite", first_suit_value)
        ut.append_child_elem(init_suit, second_suit)
        test_case = create_test_case(doc, test_msg)
        ut.append_child_elem(second_suit, test_case)
        return
    append_in_suit = ut.get_element_by_attr_name(init_suit, "testsuite", second_suit_value)
    test_case = create_test_case(doc, test_msg)
    ut.append_child_elem(append_in_suit, test_case)


def run_t(wb):
    try:
        # ws = ut.get_work_book('testlink1.xlsx')
        sheets = wb.sheetnames
        ws = wb[sheets[0]]
        doc = ut.create_dom()
        root_node = ut.create_element(doc, "testsuite", None, None)
        ut.append_child_elem(doc, root_node)
        max_row = ut.get_max_rows(ws)
        max_col = ut.get_max_col(ws)
        init_suit_value = ut.get_cell_value(ws, 2, 1)
        init_suit_value = str(init_suit_value).strip()
        init_second_suits = list()
        # print(init_suit_value)
        init_suit = ut.create_element(doc, "testsuite", None, init_suit_value)
        # print(init_suit)
        ut.append_child_elem(root_node, init_suit)
        for i in range(2, max_row + 1):
            values = ut.get_row_value(ws, i, max_col)
            first_suit_value = ut.trim_space(values[0])
            second_suit_value = ut.trim_space(values[1])
            test_msg = values[2:]
            test_msg = ut.fill_None_value(test_msg)
            if first_suit_value == init_suit_value:
                add_test(doc, second_suit_value, test_msg, init_second_suits, init_suit)
            else:
                init_suit_value = first_suit_value
                new_suit = ut.create_element(doc, "testsuite", None, first_suit_value)
                init_suit = new_suit
                ut.append_child_elem(root_node, init_suit)
                init_second_suits = list()

                add_test(doc, second_suit_value, test_msg, init_second_suits, init_suit)
        file_name = str(time.time()) + "testCase.xml"
        ut.write_xml(doc, file_name)
        print("transform testCase %d items" % (max_row-1))
        return "转化完成,请查看当前目录下的testCase.xml,共转化"+ str(max_row-1)+"条用例"
    except Exception as e:
        print(e)
        return "异常"+ str(e)


if __name__ == '__main__':
    run_t()
