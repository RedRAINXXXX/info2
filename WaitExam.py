from selenium import webdriver
import pandas as pd
import time
import os, sys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

try:
	y0 = int(input("网站起始位置：")) - 1
	x0 = int(input("公司起始位置：")) - 1
	path="诚信信息查询"
	if not os.path.exists(path):
		os.mkdir(path)
		print("总目录已创建")

	source=pd.read_excel('公司列表.xlsx',sheet_name=[0,1],header=0)#读取需要搜索的公司名称
	company_name=source[0]
	website=source[1]

	y=len(website.iloc[:,0])
	website['Open_wait_time'].astype('int')
	website['Search_wait_time'].astype('int')
	website['Valid'].astype('int')
	company_name['Valid'].astype('int')

	option = Options()
	option.add_argument('headless')

	driver = webdriver.Firefox(options=option)
	# driver = webdriver.Firefox()
	driver.implicitly_wait(30)
	driver.maximize_window()  # 窗口最大化

	scale = "0.9"

	for no in range(y0,y):
		# if no != 6 :
		# 	continue
		path0="诚信信息查询//"+website.iloc[no,0]

		if not os.path.exists(path0):
			os.mkdir(path0)
			print(website.iloc[no,0]+"子目录已创建")

		if website.iloc[no, 4] == 0:
			continue

		try:
			driver.get(website.iloc[no, 1])  # 打开网址
		except Exception as errmsg:
			print("Open website {} failure".format(website.iloc[no, 1]))
			print(errmsg)
		time.sleep(website.iloc[no, 2])
		# driver.execute_script(
		# 	"document.body.style.cssText = document.body.style.cssText + '; -moz-transform:scale("+scale+");-moz-transform-origin:top; ';")

		#no == 0
		if website.iloc[no,0]=='信用中国-失信被执行人':
			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					# className = driver.find_element_by_id(website.iloc[no,2])#使用class="##"定位搜索框
					className = driver.find_element_by_xpath("//*[@id='publishPeopleCheck']/div[1]/div/input[1]")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					className.send_keys(Keys.ENTER)
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+100
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no,0],x))
					print(errmsg)
		# no == 1
		elif website.iloc[no,0]=='信用中国-重大税收违法案件':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_xpath('//*[@id="taxCheck"]/div[1]/div/input')
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					className.send_keys(Keys.ENTER)
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+100
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 2
		elif website.iloc[no,0]=='信用中国-安全生产领域失信':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_xpath('//*[@id="flSearch"]')
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					className.send_keys(Keys.ENTER)
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+100
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 3
		elif website.iloc[no,0]=='信用中国-涉金融领域严重失信':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_xpath('//*[@id="flSearch"]')
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					className.send_keys(Keys.ENTER)
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+100
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 4
		elif website.iloc[no,0]=='信用中国-统计领域严重失信':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_xpath('//*[@id="flSearch"]')
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					className.send_keys(Keys.ENTER)
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+100
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 5
		elif website.iloc[no,0]=='信用中国-涉电力领域失信':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_xpath('//*[@id="flSearch"]')
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					className.send_keys(Keys.ENTER)
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+100
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 6
		elif website.iloc[no,0]=='中华人民共和国生态环境部':

			for x,c_valid,index in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1],range(x0,len(company_name.iloc[x0:,0]))):
				if c_valid == 0:
					continue
				try:
					if index != 0:
						try:
							driver.get(website.iloc[no, 1])  # 打开网址
							time.sleep(website.iloc[no, 2])
						except Exception as errmsg:
							print("Open website {} failure".format(website.iloc[no, 1]))
							print(errmsg)

					order = driver.find_elements_by_xpath("//*[@id='orderby']")[1].click()
					begin_year = Select(driver.find_element_by_xpath("//*[@id='s_time_1']")).select_by_value("1987")
					begin_month = Select(driver.find_element_by_xpath("//*[@id='s_time_2']")).select_by_visible_text("1")
					end_year = Select(driver.find_element_by_xpath("//*[@id='e_time_1']")).select_by_value("2021")
					end_month = Select(driver.find_element_by_xpath("//*[@id='e_time_2']")).select_by_visible_text("12")

					className = driver.find_element_by_id("orsen")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_xpath("/html/body/div/div[2]/form/div[5]/a").click()
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+120
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 7
		elif website.iloc[no,0]=='工业和信息化部网站':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_id("q")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_id("one1").click()
					time.sleep(1)
					order = driver.find_element_by_xpath("//*[@id='od-box']/div[1]")
					ActionChains(driver).move_to_element(order).perform()
					driver.find_element_by_xpath("//*[@id='od-box']/div[2]/div[2]").click()
					#button=driver.find_element_by_id('query_btn')
					#ActionChains(driver).click(button).perform()
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 8
		elif website.iloc[no,0]=='中国保险监督管理委员会':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_id("keywords")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					className.send_keys(Keys.ENTER)
					#button=driver.find_element_by_id('query_btn')
					#ActionChains(driver).click(button).perform()
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 9
		elif website.iloc[no,0]=='中国保险行业协会':

			for x,c_valid,index in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1],range(x0,len(company_name.iloc[x0:,0]))):
				if c_valid == 0:
					continue
				try:
					if index != 0:
						try:
							driver.get(website.iloc[no, 1])  # 打开网址
							time.sleep(website.iloc[no, 2])
						except Exception as errmsg:
							print("Open website {} failure".format(website.iloc[no, 1]))
							print(errmsg)

					driver.find_element_by_xpath("//*[@id='search-form']/div[1]/div/div/div[3]/div/span").click()
					time.sleep(1)
					searchBox = driver.find_element_by_id("q")
					searchBox.clear()#清除内容
					searchBox.send_keys(x) #输入搜索公司名称
					time.sleep(1.5)
					#点击空白
					ActionChains(driver).move_by_offset( 1,  1).click().perform()
					time.sleep(1)
					ActionChains(driver).move_by_offset(-1, -1).perform()
					className = driver.find_element_by_id("pq")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_xpath('//*[@id="search-form"]/div[1]/div/div/div[3]/div/input').click()
					#button=driver.find_element_by_id('query_btn')
					#ActionChains(driver).click(button).perform()
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+100
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
			pass
		# no == 10
		elif website.iloc[no,0]=='中华人民共和国商务部':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_id("input2")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_xpath('/html/body/div[3]/table/tbody/tr[1]/td/div[1]/button').click()
					time.sleep(website.iloc[no, 3])
					handles = driver.window_handles
					driver.switch_to.window(handles[1])
					#button=driver.find_element_by_id('query_btn')
					#ActionChains(driver).click(button).perform()
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+100
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.close()
					driver.switch_to.window(handles[0])
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					handles = driver.window_handles
					if len(handles) > 1:
						driver.switch_to.window(handles[1])
						driver.close()
						driver.switch_to.window(handles[0])
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 11
		elif website.iloc[no,0]=='中华人民共和国自然资源部':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_id("searchText")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_xpath('//*[@id="container"]/div[3]/div/div[2]/form/input[10]').click()
					time.sleep(website.iloc[no, 3])
					handles = driver.window_handles
					driver.switch_to.window(handles[1])
					#button=driver.find_element_by_id('query_btn')
					#ActionChains(driver).click(button).perform()
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+50
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.close()
					driver.switch_to.window(handles[0])
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					handles = driver.window_handles
					if len(handles) > 1:
						driver.switch_to.window(handles[1])
						driver.close()
						driver.switch_to.window(handles[0])
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
			pass
		# no == 12
		elif website.iloc[no,0]=='国家市场监督管理总局':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_id("qt")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_xpath('//*[@id="form1search"]/div[2]/input').click()
					time.sleep(website.iloc[no, 3])
					handles = driver.window_handles
					driver.switch_to.window(handles[1])
					ActionChains(driver).click(driver.find_element_by_xpath('/html/body/div[2]/div/div[1]/div[2]/div[3]')).perform()
					time.sleep(1)
					ActionChains(driver).click(
						driver.find_element_by_xpath('//*[@id="orderWay_downSelect"]/li[2]')).perform()

					#button=driver.find_element_by_id('query_btn')
					#ActionChains(driver).click(button).perform()

					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+50
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.close()
					driver.switch_to.window(handles[0])
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					handles = driver.window_handles
					if len(handles) > 1:
						driver.switch_to.window(handles[1])
						driver.close()
						driver.switch_to.window(handles[0])
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 13
		elif website.iloc[no,0]=='中华人民共和国财政部':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_id("andsen")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_xpath('//*[@id="searchform"]/div/a/img').click()
					time.sleep(website.iloc[no, 3])
					handles = driver.window_handles
					driver.switch_to.window(handles[1])
					#button=driver.find_element_by_id('query_btn')
					#ActionChains(driver).click(button).perform()
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+50
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.close()
					driver.switch_to.window(handles[0])
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					handles = driver.window_handles
					if len(handles) > 1:
						driver.switch_to.window(handles[1])
						driver.close()
						driver.switch_to.window(handles[0])
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 14
		elif website.iloc[no,0]=='国家发展和改革委员会':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_id("qt")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/dl/dd/form/input[5]').click()
					time.sleep(website.iloc[no, 3])
					handles = driver.window_handles
					driver.switch_to.window(handles[1])
					#button=driver.find_element_by_id('query_btn')
					#ActionChains(driver).click(button).perform()
					driver.find_element_by_xpath('/html/body/div[6]/div[2]/div[1]/div[1]/a[1]').click()
					time.sleep(website.iloc[no, 3])
					driver.find_element_by_xpath('/html/body/div[6]/div[2]/div[1]/div[2]/a[1]').click()
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+50
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.close()
					driver.switch_to.window(handles[0])
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					handles = driver.window_handles
					if len(handles) > 1:
						driver.switch_to.window(handles[1])
						driver.close()
						driver.switch_to.window(handles[0])
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 15
		elif website.iloc[no,0]=='中华人民共和国农业农村部':

			for x,c_valid in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1]):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_xpath('/html/body/div[4]/div[1]/div/form/label/input[1]')
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_xpath('//*[@id="search_btn"]').click()
					time.sleep(website.iloc[no, 3])
					handles = driver.window_handles
					driver.switch_to.window(handles[1])
					#button=driver.find_element_by_id('query_btn')
					#ActionChains(driver).click(button).perform()
					driver.find_element_by_xpath('/html/body/div[1]/div[1]/form/a').click()
					time.sleep(website.iloc[no, 3])
					className2 = driver.find_element_by_xpath('//*[@id="k2"]')
					className2.clear()#清除内容
					className2.send_keys(x) #输入搜索公司名称
					driver.find_element_by_xpath('//*[@id="searchform"]/ul/li[5]/div/input[1]').click()
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+50
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.close()
					driver.switch_to.window(handles[0])
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					handles = driver.window_handles
					if len(handles) > 1:
						driver.switch_to.window(handles[1])
						driver.close()
						driver.switch_to.window(handles[0])
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
		# no == 16
		elif website.iloc[no,0]=='住房和城乡建筑部':

			for x,c_valid,index in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1],range(x0,len(company_name.iloc[x0:,0]))):
				if c_valid == 0:
					continue
				try:
					if index != 0:
						try:
							driver.get(website.iloc[no, 1])  # 打开网址
							time.sleep(website.iloc[no, 2])
						except Exception as errmsg:
							print("Open website {} failure".format(website.iloc[no, 1]))
							print(errmsg)

					className = driver.find_element_by_id('ukf')
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_xpath('/html/body/div[1]/main/section/div/div[7]/div/label[1]/input').click()
					time.sleep(1)
					driver.find_element_by_xpath('//*[@id="advanced-search"]').click()
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '.png'
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+50
					driver.set_window_size(width, height)
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)
			pass
		# no == 17
		elif website.iloc[no,0]=='百度':

			for x,c_valid,index in zip(company_name.iloc[x0:,0],company_name.iloc[x0:,1],range(x0,len(company_name.iloc[x0:,0]))):
				if c_valid == 0:
					continue
				try:
					className = driver.find_element_by_id("kw")
					className.clear()#清除内容
					className.send_keys(x) #输入搜索公司名称
					driver.find_element_by_id("su").click()
					time.sleep(website.iloc[no, 3])
					if index == 0:
						driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/a[1]').click()
						time.sleep(website.iloc[no, 3])
					width = driver.execute_script("return document.documentElement.scrollWidth")
					height = driver.execute_script("return document.documentElement.scrollHeight")+100
					driver.set_window_size(width, height)
					filename = path0 + '//' + x + '_1.png'
					driver.save_screenshot(filename)
					driver.find_element_by_xpath('//*[@id="page"]/div/a[1]/span[2]').click()
					time.sleep(website.iloc[no, 3])
					filename = path0 + '//' + x + '_2.png'
					driver.save_screenshot(filename)
					driver.maximize_window()
					print("查询网站:{} 查询公司:{} 截图成功".format(website.iloc[no, 0], x))
				except Exception as errmsg:
					print("Error location: {}  {} ".format(website.iloc[no, 0], x))
					print(errmsg)

	driver.quit()
except Exception as errmsg:
	print(errmsg)

input('Press <Enter> to quit.')