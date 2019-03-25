from Util.Excel import  *
from Config.ProjVar import *
from Action.WebElementAction import *
from Util.ObjectMap import *
from Action.ResetTestCaseFileResult import *

def execute_test_cases(test_data_file_path):
    clear_all_executed_info(test_data_file_path)
    test_data_wb = ParseExcel(test_data_file_path)
    test_data_wb.set_sheet_by_name(test_case_sheet)
    col_cells = test_data_wb.get_col(test_case_is_executed_col_no)
    success_step_num = 0
    for id,i in enumerate(range(1,len(col_cells))):
        print(id+2,col_cells[i].value)
        if col_cells[i].value.lower()=="y":
            test_step_sheet = test_data_wb.get_cell_value(id+2,3)
            print(test_step_sheet)
            test_data_wb.set_sheet_by_name(test_step_sheet)
            max_step_row_no = test_data_wb.get_max_row()
            for i in range(2,max_step_row_no+1):
                step_row = test_data_wb.get_row(i)
                action = step_row[test_step_action_col_no-1].value
                locate_method = step_row[test_step_locate_type_col_no-1].value
                locate_expression = step_row[test_step_locate_expression_col_no - 1].value
                if action: action=action.strip()
                if locate_method: locate_method = locate_method.strip()
                if locate_expression: locate_expression = locate_expression.strip()
                if locate_expression is not None  and "Page." in locate_expression:
                    locate_method,locate_expression = ObjectMap(object_map_file_path).get_locatemethod_and_locateexpression(
                        locate_method,locate_expression
                    )
                value = step_row[test_step_value_col_no-1].value
                print(action,locate_method,locate_expression
                      ,value)
                #eval("open_browser('ie')")
                if  action is not None and locate_method is None \
                        and locate_expression is None and value is not None:
                    command = "%s('%s')" %(action,value)
                elif action is not None and locate_method is not None \
                        and locate_expression is not None and value is not None:
                    command = "%s('%s','%s','%s')" %(action,locate_method,
                                                     locate_expression,value)
                elif action is not None and locate_method is not None \
                     and locate_expression is not None and value is None:
                    command = "%s('%s','%s')" % (action, locate_method,
                                                  locate_expression)
                elif action is not None and locate_method is None \
                     and locate_expression is None and value is None:
                    command = "%s()" % (action)
                print(command)
                try:
                    return_value  = eval(command)
                    test_data_wb.write_cell(i,test_step_executed_result_col_no,"pass")
                    success_step_num+=1
                    if "capture_pic" in command:
                        test_data_wb.write_cell(
                            i, test_step_executed_capture_pic_path_col_no, return_value)
                except AssertionError as e:
                    info(command + "\n" + "断言失败：\n" + traceback.format_exc())
                    test_data_wb.write_cell(i, test_step_executed_result_col_no, "fail")
                    pic_path = capture_pic()
                    test_data_wb.write_cell(
                        i, test_step_executed_capture_pic_path_col_no, pic_path)
                    test_data_wb.write_cell(
                        i, test_step_executed_exception_info_col_no, traceback.format_exc())

                except Exception as e:
                    capture_pic()
                    info(command+"\n"+e+"\n"+traceback.format_exc())
                    test_data_wb.write_cell(i, test_step_executed_result_col_no, "fail")
                    test_data_wb.write_cell(
                        i, test_step_executed_capture_pic_path_col_no, pic_path)
                    test_data_wb.write_cell(
                        i, test_step_executed_exception_info_col_no, traceback.format_exc())
                test_data_wb.write_cell(i,test_step_executed_time_col_no, get_current_datetime())
            if success_step_num == max_step_row_no-1:
                test_data_wb.set_sheet_by_name(test_case_sheet)
                test_data_wb.write_cell(id+2,test_case_executed_result_col_no,"pass")
            else:
                test_data_wb.set_sheet_by_name(test_case_sheet)
                test_data_wb.write_cell(id + 2, test_case_executed_result_col_no, "fail")
            test_data_wb.write_cell(id + 2, test_case_executed_time_col_no, get_current_datetime())



if __name__ == "__main__":
    execute_test_cases(test_data_file)
    #print(execute_test_case(test_data_file+"||搜狗"))