from selenium import webdriver
from Config.ProjVar import *
import time
from Util.Dir import *
from Util.GetTime import *
from Util.WaitUtil import  *
from Util.Log import *
from Util.Excel import *
from Util.ObjectMap import *
from Action.ResetTestCaseFileResult import *
driver = None

def capture_pic():
    global driver
    try:
        pic_dir_path = os.path.join(proj_path, "capture_pics")
        pic_dir_path = os.path.join(pic_dir_path,make_date_dir(pic_dir_path))
        pic_path = os.path.join(pic_dir_path, get_current_time() + ".png")
        result = driver.get_screenshot_as_file(os.path.join(pic_path))
        print(result)
    except IOError as e:
        info(e)
        traceback.print_exc()
    except Exception as e:
        info(e)
        traceback.print_exc()

def open_browser(browser_name):
    global driver
    if browser_name.lower() =="ie":
        driver = webdriver.Ie(executable_path=ie_driver_path)
    elif browser_name.lower() =="chrome":
        driver = webdriver.Chrome(executable_path=chrome_driver_path)
    else:
        driver = webdriver.Firefox(executable_path=firefox_driver_path)

def visit(url):
    global driver
    try:
        driver.get(url)
    except:
        info("%s can not be visited!" %url)
        raise Exception("网址%s无法访问" %url)

def input(locate_method,locate_expression,content):
    global driver
    try:
        element = WaitUtil(driver).visibleOfElement(
            locate_method,locate_expression)
        element.send_keys(content)
    except:
        print("输入内容出现了异常")
        info("webelement:%s->%s  operate fail" %(
            locate_method,locate_expression))

def sleep(duration):
    time.sleep(int(duration))

def click(locate_method,locate_expression):
    global driver
    try:
        element = WaitUtil(driver).visibleOfElement(
            locate_method,locate_expression)
        element.click()
    except:
        print("输入内容出现了异常")
        info("webelement:%s->%s  operate fail" %(
            locate_method,locate_expression))

def assert_word(word):
    global driver
    assert word in driver.page_source

def quit():
    global driver
    driver.quit()

def execute_test_case(test_data_file_path_and_sheet_name):
    test_data_file_path, sheet_name = test_data_file_path_and_sheet_name.split("||")
    clear_all_executed_info(test_data_file_path)
    test_data_wb = ParseExcel(test_data_file_path)
    test_data_wb.set_sheet_by_name(sheet_name)
    success_step_num = 0
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
        return True
    else:
        return False

if __name__ =="__main__":
    try:
        open_browser("chrome")
        visit("http://www.sogou.com")
        capture_pic()
    except Exception as e:
        print("测试用例执行失败了！原因是:")
        traceback.print_exc()
    finally:
        quit()


