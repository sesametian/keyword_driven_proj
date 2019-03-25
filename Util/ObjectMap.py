#encoding=utf-8
from selenium.webdriver.support.ui import WebDriverWait
from Util.GetConfig import *
import configparser
import os
from Config.ProjVar import *
from selenium import webdriver

class ObjectMap(object):
    def __init__(self,config_file_path):
        self.config_file_path = config_file_path
        self.cf = Config(self.config_file_path)

    def get_locatemethod_and_locateexpression(self,webSiteName,elementName):
        locators = self.cf.get_option(webSiteName, elementName).split(">")
        return locators

    def getElementObject(self, driver, webSiteName, elementName):
        try:
            locators = self.cf.get_option(webSiteName, elementName).split(">")
            # 得到定位方式
            locatorMethod = locators[0]
            # 得到定位表达式
            locatorExpression = locators[1]
            #print (locatorMethod, locatorExpression)
            # 通过显示等待方式获取页面元素
            element = WebDriverWait(driver, 10).until(lambda x: \
                    x.find_element(locatorMethod, locatorExpression))
        except Exception as e:
            raise e
        else:
            # 当页面元素被找到后，将该页面元素对象返回给调用者
            return element

if __name__ == '__main__':
    driver = webdriver.Chrome(executable_path = "e:\\chromedriver")
    url = "http://www.baidu.com"
    driver.get(url)
    objmap = ObjectMap(object_map_file_path)
    print (objmap.getElementObject(driver,"baidu","SearchPage.InputBox"))
    print (objmap.getElementObject(driver,"baidu","SearchPage.SubmitButton"))

