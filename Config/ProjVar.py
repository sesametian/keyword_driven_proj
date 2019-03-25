import os

proj_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
conf_path = os.path.join(proj_path,"config","Logger.conf")

ie_driver_path = "e:\\IEDriverServer"
chrome_driver_path = "e:\\chromeDriver"
firefox_driver_path = "e:\\geckoDriver"

test_data_file = os.path.join(proj_path,"TestData",\
"测试用例.xlsx")
test_case_sheet = "测试用例"
test_case_test_step_sheet_name_col_no = 3
test_case_is_executed_col_no = 4
test_case_executed_time_col_no = 5
test_case_executed_result_col_no = 6

test_step_executed_exception_info_col_no = 8
test_step_executed_capture_pic_path_col_no = 9
test_step_action_col_no = 2
test_step_locate_type_col_no = 3
test_step_locate_expression_col_no = 4
test_step_value_col_no = 5
test_step_executed_time_col_no = 6
test_step_executed_result_col_no = 7


object_map_file_path = os.path.join(proj_path,"testdata","ObjectDeposit.ini")
if __name__ =="__main__":
    print(conf_path)
