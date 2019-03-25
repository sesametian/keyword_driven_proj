from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import traceback
from selenium.common.exceptions import TimeoutException

class WaitUtil(object):

    def __init__(self,driver):
        self.driver = driver
        self.wait =  WebDriverWait(self.driver,5)
        self.locate_method = {"id":By.ID,
                             "name":By.NAME,
                              "link_text":By.LINK_TEXT,
                              "partial_link_text":By.PARTIAL_LINK_TEXT,
                              "xpath":By.XPATH
                              }


    def presenceOfElement(self,locate_method,locate_expression):
        try:
            element = self.wait.until \
                (lambda x: x.find_element(self.locate_method[locate_method]
                                          ,locate_expression))
            return element
        except TimeoutException:
            traceback.print_exc()
            raise TimeoutException

    def visibleOfElement(self,locate_method,locate_expression):
        try:
            element = self.wait.until(\
                    EC.visibility_of_element_located((
                        self.locate_method[locate_method]
                        , locate_expression)))
            return element
        except TimeoutException:
            traceback.print_exc()
            raise TimeoutException

    def switchToFrame(self,locate_method,locate_expression):
        try:
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it((self.locate_method[locate_method]
                            , locate_expression)))
        except Exception as e:
            raise e

if __name__ == "__main__":
    #driver = webdriver.Ie(executable_path="e:\\IEDriverServer")
    driver = webdriver.Firefox(executable_path="e:\\geckodriver")
    wait_object = WaitUtil(driver)
    driver.get("http://mail.126.com")
    try:
        element = wait_object.switchToFrame("xpath","//iframe[contains(@id,'x-URS-iframe')]")
        import time
        time.sleep(3)
    except TimeoutException:
        print("元素未定位！")
