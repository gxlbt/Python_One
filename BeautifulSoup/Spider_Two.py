#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019-2-15 9:44
# @File    : Spider_One  根据不同页码抓取
# @Software: PyCharm

from selenium import webdriver
import time
import xlwt

browser = webdriver.Chrome()
count = 0
try:
    new_workbook = xlwt.Workbook()  # 创建工作簿
    new_sheet = new_workbook.add_sheet("Sheet1")  # 创建表
    new_sheet.write(0, 0, "品牌名称")
    new_sheet.write(0, 1, "申报型号")
    new_sheet.write(0, 2, "外形尺寸")
    new_sheet.write(0, 3, "整车质量")
    new_sheet.write(0, 4, "最高车速")
    new_sheet.write(0, 5, "企业名称")
    new_sheet.write(0, 6, "前轮宽")
    new_sheet.write(0, 7, "后轮宽")
    new_sheet.write(0, 8, "前后轮中心距")
    new_sheet.write(0, 9, "空车质量")
    new_sheet.write(0, 10, "行驶距离")
    new_sheet.write(0, 11, "检验结论")
    new_sheet.write(0, 12, "电池类型")
    new_sheet.write(0, 13, "图片地址")
    new_sheet.write(0, 14, "批次")

    browser.get('http://bs.gxqts.gov.cn:8080/gxwsdt/wsbsdt/PublicSearch/SearchResult2.jspx?pageCode=008&pageText=%E7'
                '%94%B5%E5%8A%A8%E8%87%AA%E8%A1%8C%E8%BD%A6%E6%B3%A8%E5%86%8C%E7%99%BB%E8%AE%B0%E7%9B%AE%E5%BD%95')
    now_handle = browser.current_window_handle
    browser.find_element_by_xpath('//*[@id="txt_jump"]').send_keys('970')
    time.sleep(1)
    browser.find_element_by_xpath('//*[@id="MPageID"]/a[10]').click()

    # //*[@id="MPageID"]/a[7]

    for i in range(970, 974):
        if i > 970:
            count = count + 10
            time.sleep(2)
            # browser.find_element_by_xpath('//*[@id="MPageID"]/a[9]').click()
            browser.find_element_by_xpath('//*[@id="MPageID"]/a[7]').click()
        tags = browser.find_elements_by_css_selector('#tr > td:nth-child(13) > a')
        for j in range(len(tags)):
            time.sleep(1)
            pc = browser.find_element_by_xpath('//*[@id="tr"]/td[2]').text
            time.sleep(2)
            tags[j].click()
            if i > 1:
                browser.switch_to.frame('layui-layer-iframe' + str(count + (j + 1)))
            else:
                browser.switch_to.frame('layui-layer-iframe' + str(j + 1))
            s = browser.find_element_by_class_name('main')
            time.sleep(3)
            scqymc = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[1]/td[2]').text  # 生产企业名称/html/body
            tp = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[1]/td[3]').find_element_by_xpath(
                '//*[@id="splct"]').get_attribute('src')
            sbpp = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[2]/td[2]').text  # 商标品牌
            xh = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[2]/td[4]').text  # 型号
            wxcc = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[3]/td[2]').text  # 外形尺寸
            qlk = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[3]/td[4]').text  # 前轮宽
            hlk = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[4]/td[2]').text  # 后轮宽
            qhlzxj = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[4]/td[4]').text  # 前后轮中心距
            zczl = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[5]/td[2]').text  # 整车质量
            kczl = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[5]/td[4]').text  # 空车质量
            xsjl = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[6]/td[2]').text  # 脚踏行驶距离
            zgcs = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[6]/td[4]').text  # 最高车速
            jyjl = s.find_element_by_xpath('//*[@id="spelimtestcon"]').text  # 检验结论
            dclx = s.find_element_by_xpath('/html/body/div[1]/div/table/tbody/tr[7]/td[4]').text  # 电池类型
            # print(pc + ',' + scqymc + ',' + tp + ',' + sbpp + ',' + xh + ',' + wxcc + ',' + qlk + ',' + hlk + ','
            #       + qhlzxj + ',' + zczl + ',' + kczl + ',' + xsjl + ',' + zgcs + '' + jyjl + ',' + dclx)
            time.sleep(3)
            browser.find_element_by_xpath('/html/body/div[1]/div/button').click()
            if i <= 1:
                new_sheet.write(j+1, 0, sbpp)
                new_sheet.write(j+1, 1, xh)
                new_sheet.write(j+1, 2, wxcc)
                new_sheet.write(j+1, 3, zczl)
                new_sheet.write(j+1, 4, zgcs)
                new_sheet.write(j+1, 5, scqymc)
                new_sheet.write(j+1, 6, qlk)
                new_sheet.write(j+1, 7, hlk)
                new_sheet.write(j+1, 8, qhlzxj)
                new_sheet.write(j+1, 9, kczl)
                new_sheet.write(j+1, 10, xsjl)
                new_sheet.write(j+1, 11, jyjl)
                new_sheet.write(j+1, 12, dclx)
                new_sheet.write(j+1, 13, tp)
                new_sheet.write(j+1, 14, pc)
            else:
                new_sheet.write(count + (j + 1), 0, sbpp)
                new_sheet.write(count + (j + 1), 1, xh)
                new_sheet.write(count + (j + 1), 2, wxcc)
                new_sheet.write(count + (j + 1), 3, zczl)
                new_sheet.write(count + (j + 1), 4, zgcs)
                new_sheet.write(count + (j + 1), 5, scqymc)
                new_sheet.write(count + (j + 1), 6, qlk)
                new_sheet.write(count + (j + 1), 7, hlk)
                new_sheet.write(count + (j + 1), 8, qhlzxj)
                new_sheet.write(count + (j + 1), 9, kczl)
                new_sheet.write(count + (j + 1), 10, xsjl)
                new_sheet.write(count + (j + 1), 11, jyjl)
                new_sheet.write(count + (j + 1), 12, dclx)
                new_sheet.write(count + (j + 1), 13, tp)
                new_sheet.write(count + (j + 1), 14, pc)
finally:
    new_workbook.save(r"test19.xls")  # 保存此表
    time.sleep(3)
    browser.close()
