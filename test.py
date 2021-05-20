from bs4 import BeautifulSoup
from lxml import etree
from selenium import webdriver
from lxml import html
import xlsxwriter
from mistletoe import markdown
from html2text import HTML2Text
import time
from selenium.webdriver.common.action_chains import ActionChains

link = "https://www.aliexpress.com/item/1005001380809969.html"

driver = webdriver.Chrome("/Users/ballmerpeak/Downloads/chromedriver")
driver.get(link)
time.sleep(2)
element_to_click = driver.find_element_by_xpath("//span[@class=\"SkuValueButtonItem-module_buttonItem__3x6ux\"")
actions = ActionChains(driver)
actions.click(element_to_click)
actions.perform()

time.sleep(2)
element_to_click = driver.find_element_by_xpath("//span[@class=\"product-shipping-info black-link\"]")
actions = ActionChains(driver)
actions.click(element_to_click)
actions.perform()

