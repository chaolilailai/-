import requests
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from time import sleep
import csv
import pandas as pd
import os
import json

driver = webdriver.Chrome(ChromeDriverManager().install())
url = 'https://aiqicha.baidu.com/'
driver.get(url)
# driver.maximize_window()
main_page = driver.current_window_handle  # 保存当前窗口句柄
driver.find_element_by_xpath(
    '/html/body/div[1]/header/div/div[1]/div[1]/span[1]').click()
driver.find_element_by_xpath(
    '//*[@id="TANGRAM__PSP_4__userName"]').send_keys('芳菲菲ws')
driver.find_element_by_xpath(
    '//*[@id="TANGRAM__PSP_4__password"]').send_keys('Barca1899')
driver.find_element_by_xpath(
    '//*[@id="TANGRAM__PSP_4__submit"]').click()

while True:
    search_way = input('请问要怎么搜索？ 1为直接搜索，2为高级搜索, 其他为退出')
    if search_way == '1':
        search_content = input('请输入查询内容：')
        driver.find_element_by_xpath(
            '//*[@id="aqc-search-input"]').send_keys(search_content)
        driver.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div[2]/div[2]/button').click()  # 点击“查一下”
        driver.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div[1]/div[2]/div[2]/div/div[1]/div[2]/div/h3/a').click()  # 点击搜索结果第一家
        sleep(2)

        # 切换窗口
        for handle in driver.window_handles:
            if handle != main_page:
                major_company = handle
                driver.switch_to.window(major_company)
                break

        company_name = driver.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[2]/h2').text
        legal_representative = driver.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[4]/div[1]/a').text
        registered_capital = driver.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[4]/div[2]/span').text
        foundation_date = driver.find_element_by_xpath(
            '//*[@id="basic-business"]/table/tbody/tr[6]/td[4]').text
        status = driver.find_element_by_xpath(
            '//*[@id="basic-business"]/table/tbody/tr[1]/td[4]').text
        company_type = driver.find_element_by_xpath(
            '//*[@id="basic-business"]/table/tbody/tr[7]/td[2]').text
        unified_social_credit_code = driver.find_element_by_xpath(
            '//*[@id="basic-business"]/table/tbody/tr[4]/td[2]').text
        phone_number = driver.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[5]/div[1]/div[1]/div/span[2]/span').text
        which_industry = driver.find_element_by_xpath(
            '//*[@id="basic-business"]/table/tbody/tr[3]/td[4]').text
        address = driver.find_element_by_xpath(
            '//*[@id="basic-business"]/table/tbody/tr[10]/td[2]').text.split(' ')[0]

        # 通过百度地图api获取行政区划
        my_baidu_key = 'egQBavZa6yChWgM2l7LaWc1gRez87qYT'
        lonlat_resp = json.loads(requests.get(
            f'https://api.map.baidu.com/geocoding/v3/?address={address}&output=json&ak={my_baidu_key}').text)['result'][
            'location']
        lonlat = str(lonlat_resp['lat']) + ',' + str(lonlat_resp['lng'])
        address_resp = json.loads(requests.get(
            f'https://api.map.baidu.com/reverse_geocoding/v3/?ak={my_baidu_key}&output=json&location={lonlat}').text)
        from_which_province = address_resp['result']['addressComponent']['province']
        from_which_city = address_resp['result']['addressComponent']['city']
        from_which_district = address_resp['result']['addressComponent']['district']

        website = driver.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[5]/div[2]/div[1]/div/a').text
        email = driver.find_element_by_xpath(
            '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[5]/div[1]/div[2]/div/a').text
        what_they_do = driver.find_element_by_xpath(
            '//*[@id="basic-business"]/table/tbody/tr[11]/td[2]/div').text
        temp3 = driver.find_element_by_xpath(
            '//*[@id="basic-business"]/table/tbody/tr[9]/td[2]').text
        number_of_insurance_participants = temp3.split(' ')[0]

        # 存储数据
        with open('公司信息（直接搜索）.csv', 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['企业名称', '法定代表人', '注册资本', '成立日期', '经营状态', '所属省份', '所属市区', '所属区县', '公司类型',
                             '统一社会信用代码', '联系电话', '所属行业', '地址', '网址', '邮箱', '经营范围', '参保人数'])
            writer.writerow(
                [company_name, legal_representative, registered_capital, foundation_date, status, from_which_province,
                 from_which_city, from_which_district, company_type, unified_social_credit_code,
                 phone_number, which_industry, address, website, email, what_they_do, number_of_insurance_participants])

        driver.switch_to.window(main_page)
        sleep(2)

        print('搜索完毕，数据存储完毕')

    elif search_way == '2':
        # 提前写好csv文件的表头
        if not os.path.exists('公司信息（高级搜索）.csv'):
            with open('公司信息（高级搜索）.csv', 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['企业名称', '法定代表人', '注册资本', '成立日期', '经营状态', '所属省份', '所属市区', '所属区县', '公司类型',
                                 '统一社会信用代码', '联系电话', '所属行业', '地址', '网址', '邮箱', '经营范围', '参保人数'])

        industries = {'农、林、牧、渔业': 1, '采矿业': 2, '制造业': 3, '电力、热力、燃气及水生产和供应业': 4, '建筑业': 5, '批发和零售业': 6, '交通运输、仓储和邮政业': 7,
                      '住宿和餐饮业': 8, '信息传输、软件和信息技术服务业': 9, '金融业': 10, '房地产业': 11, '租赁和商务服务业': 12, '科学研究和技术服务业': 13,
                      '水利、环境和公共设施管理业': 14, '居民服务、修理和其他服务': 15, '教育': 16, '卫生和社会工作': 17, '文化、体育和娱乐业': 18,
                      '公共管理、社会保障和社会组织': 19, '国际组织': 20}

        provinces = {'北京': 1, '天津': 2, '河北': 3, '山西': 4, '内蒙': 5, '辽宁': 6, '吉林': 7, '黑龙江': 8, '上海': 9, '江苏': 10,
                     '浙江': 11, '安徽': 12, '福建': 13, '江西': 14, '山东': 15, '河南': 16, '湖北': 17, '湖南': 18, '广东': 19, '广西': 20,
                     '海南': 21,
                     '重庆': 22, '四川': 23, '贵州': 24, '云南': 25, '西藏': 26, '陕西': 27, '甘肃': 28, '青海': 29, '宁夏': 30, '新疆': 31}

        establish_years = {'1年内': 1, '1-2年': 2, '2-3年': 3, '3-5年': 4, '5-10年': 5, '10年以上': 6}

        company_statuses = {'开业': 1, '吊销': 2, '注销': 3, '迁出': 4, '停业': 5, '撤销': 6, '解散': 7, '个体转企业': 8, '其他': 9}

        while True:
            print('行业分类选择：', '农、林、牧、渔业', '采矿业', '制造业', '电力、热力、燃气及水生产和供应业', '建筑业', '批发和零售业', '交通运输、仓储和邮政业', '住宿和餐饮业',
                  '信息传输、软件和信息技术服务业', '金融业', '房地产业', '租赁和商务服务业', '科学研究和技术服务业', '水利、环境和公共设施管理业', '居民服务、修理和其他服务', '教育',
                  '卫生和社会工作', '文化、体育和娱乐业', '公共管理、社会保障和社会组织', '国际组织')
            industry_input = input('请输入行业分类：')
            if industry_input in industries.keys():
                industry = industries[industry_input]  # 按照行业赋值
                break
            else:
                print('该行业分类不存在，请重新输入。')

        while True:
            print('省份选择：', '北京', '天津', '河北', '山西', '内蒙', '辽宁', '吉林', '黑龙江', '上海', '江苏', '浙江', '安徽', '福建', '江西', '山东',
                  '河南', '湖北', '湖南', '广东', '广西', '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃', '青海', '宁夏', '新疆')
            province_input = input('请输入地区：')
            if province_input in provinces.keys():
                province = provinces[province_input]  # 按照省份赋值
                break
            else:
                print('该地区不存在，请重新输入。')

        while True:
            print('成立年限选择：', '1年内', '1-2年', '2-3年', '3-5年', '5-10年', '10年以上')
            establish_years_input = input('请输入成立年限：')
            if establish_years_input in establish_years.keys():
                establish_year = establish_years[establish_years_input]  # 按照成立年限赋值
                break
            else:
                print('该成立年限不存在，请重新输入。')

        while True:
            print('企业状态选择：', '开业', '吊销', '注销', '迁出', '停业', '撤销', '解散', '个体转企业', '其他')
            company_status_input = input('请输入公司状态：')
            if company_status_input in company_statuses.keys():
                company_status = company_statuses[company_status_input]  # 按照企业状态赋值
                break
            else:
                print('该企业状态不存在，请重新输入。')

        driver.find_element_by_xpath(
            '/html/body/div[1]/div[2]/div[1]/div[1]/div/div[6]/a/img').click()  # 高级搜索
        driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[3]/div[2]/div/div/div[1]/div/span').click()  # 点击“全部行业”
        driver.find_element_by_xpath(
            f'/html/body/div[1]/div/div/div[3]/div[2]/div/div/div[2]/div/div[2]/div/div/div/div[1]/div[1]/ul/li[{industry}]/div/label/span/input').click()  # 点击行业选项卡
        driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[3]/div[2]/div/div/div[2]/div/div[2]/div/div/div/div[2]/button[2]/span').click()  # 点击确定
        sleep(2)
        driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[5]/div[2]/div/div/div[1]/div/span').click()  # 点击“全部地区”
        driver.find_element_by_xpath(
            f'/html/body/div[1]/div/div/div[5]/div[2]/div/div/div[2]/div/div[2]/div/div/div/div[1]/div[1]/ul/li[{province}]/div/label/span/input').click()  # 选择省份
        driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[5]/div[2]/div/div/div[2]/div/div[2]/div/div/div/div[2]/button[2]/span').click()  # 点击确定
        sleep(2)
        driver.find_element_by_xpath(
            f'/html/body/div[1]/div/div/div[7]/div[2]/div[1]/label[{establish_year}]/span/input').click()  # 选择成立年限

        driver.find_element_by_xpath(
            f'/html/body/div[1]/div/div/div[10]/div[2]/div/label[{company_status}]/span/input').click()  # 选择公司状态

        driver.find_element_by_xpath('/html/body/div[1]/div/div/div[16]/div[2]/button[3]/span').click()  # 点击查看结果
        sleep(2)

        search_number = int(input('请输入查询次数：'))

        # 读之前写好的xls文件
        with open('公司信息（高级搜索）.csv', 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            lines = [line for line in reader]
        existing_company_num = len(lines) - 1  # 减一是为了去掉表头

        for count in range(existing_company_num + 1, existing_company_num + search_number + 1):  # 根据已有数据表跳过搜索
            # 如果页码是10，那么取模后是0，但是没有第0页，所以依然取10，其余页码保持不变
            if count % 10 == 0:
                page_number = 10
            else:
                page_number = count % 10
            next_page_number = count // 10 + 1  # 一共多少页

            # 这里其实没必要切换，但是有个莫名其妙的bug，就切一下吧
            driver.switch_to.window(main_page)
            sleep(2)

            # 输入页码
            try:
                ele = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[3]/ul/div/div/input')
                ele.clear()
                ele.send_keys(next_page_number)
                ele.send_keys(Keys.ENTER)
            except NoSuchElementException:
                pass

            # 把目标元素拉动到可见区域
            js = 'var q=document.documentElement.scrollTop=0'
            driver.execute_script(js)

            # 获取目标内容
            try:
                driver.find_element_by_xpath(
                    f'/html/body/div[1]/div/div/div[3]/div[2]/div/div[{page_number}]/div[2]/div/h3/a').click()
                sleep(2)

                # 切换到最新打开的窗口
                windows = driver.window_handles
                driver.switch_to.window(windows[-1])
                sleep(2)

                company_name = driver.find_element_by_xpath(
                    '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[2]/h2').text
                print(company_name)
                legal_representative = driver.find_element_by_xpath(
                    '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[4]/div[1]/a').text
                print(legal_representative)
                registered_capital = driver.find_element_by_xpath(
                    '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[4]/div[2]/span').text
                print(registered_capital)
                foundation_date = driver.find_element_by_xpath(
                    '//*[@id="basic-business"]/table/tbody/tr[6]/td[4]').text
                print(foundation_date)
                status = driver.find_element_by_xpath(
                    '//*[@id="basic-business"]/table/tbody/tr[1]/td[4]').text
                print(status)
                company_type = driver.find_element_by_xpath(
                    '//*[@id="basic-business"]/table/tbody/tr[7]/td[2]').text
                print(company_type)
                unified_social_credit_code = driver.find_element_by_xpath(
                    '//*[@id="basic-business"]/table/tbody/tr[4]/td[2]').text
                print(unified_social_credit_code)
                phone_number = driver.find_element_by_xpath(
                    '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[5]/div[1]/div[1]/div/span[2]/span').text
                print(phone_number)
                which_industry = driver.find_element_by_xpath(
                    '//*[@id="basic-business"]/table/tbody/tr[3]/td[4]').text
                print(which_industry)
                address = driver.find_element_by_xpath(
                    '//*[@id="basic-business"]/table/tbody/tr[10]/td[2]').text.split(' ')[0]
                print(address)

                # 通过百度地图api获取行政区划
                my_baidu_key = 'egQBavZa6yChWgM2l7LaWc1gRez87qYT'
                try:
                    lonlat_resp = json.loads(requests.get(
                        f'https://api.map.baidu.com/geocoding/v3/?address={address}&output=json&ak={my_baidu_key}').text)[
                        'result']['location']
                    lonlat = str(lonlat_resp['lat']) + ',' + str(lonlat_resp['lng'])
                    address_resp = json.loads(requests.get(
                        f'https://api.map.baidu.com/reverse_geocoding/v3/?ak={my_baidu_key}&output=json&location={lonlat}').text)
                    from_which_province = address_resp['result']['addressComponent']['province']
                    from_which_city = address_resp['result']['addressComponent']['city']
                    from_which_district = address_resp['result']['addressComponent']['district']
                except KeyError:
                    from_which_province = '?'
                    from_which_city = '?'
                    from_which_district = '?'
                print(from_which_province)
                print(from_which_city)
                print(from_which_district)

                website = driver.find_element_by_xpath(
                    '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[5]/div[2]/div[1]/div/span[2]').text
                print(website)
                email = driver.find_element_by_xpath(
                    '/html/body/div[1]/div[1]/div/div[2]/div[1]/div[1]/div[2]/div[5]/div[1]/div[2]/div/a').text
                print(email)
                what_they_do = driver.find_element_by_xpath(
                    '//*[@id="basic-business"]/table/tbody/tr[11]/td[2]/div').text
                print(what_they_do)
                number_of_insurance_participants = driver.find_element_by_xpath(
                    '//*[@id="basic-business"]/table/tbody/tr[9]/td[2]').text.split(' ')[0]
                print(number_of_insurance_participants)
                print('===========================================')

                with open('公司信息（高级搜索）.csv', 'a', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow([company_name, legal_representative, registered_capital, foundation_date, status,
                                     from_which_province,
                                     from_which_city, from_which_district, company_type, unified_social_credit_code,
                                     phone_number, which_industry, address, website, email, what_they_do,
                                     number_of_insurance_participants])

                driver.close()
                driver.switch_to.window(main_page)

            except NoSuchElementException:
                pass

            sleep(2)

        print('搜索完毕，数据存储完毕')

    else:
        print('程序已退出')
        break
