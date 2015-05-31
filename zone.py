from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup as bs
import xlwt


driver = webdriver.Firefox()
driver.get("http://studzone.psgtech.edu")
driver.execute_script("javascript:__doPostBack('DgdLinks:_ctl2','links')")
wb = xlwt.Workbook()
ws = wb.add_sheet("Attendance")
row = 2
col =1

def crawl(l,h,st,row,col):

	
	
	for no in range(l,h):

		rollno=driver.find_element_by_name("TxtRollNo")
		rollno.clear()
		
		if no<10 and st=="regular":
			Rno = "13R20" + str(no)
		if no>=10 and st == "regular":
			Rno = "13R2" + str(no)
		if st=="lateral":
			Rno = "14R4" + str(no)
		
		print " Extracting info of Roll No  " + str(Rno)
	
		rollno.send_keys(Rno)
		rollno.send_keys(Keys.RETURN)
		soup = bs(driver.page_source)
		
		if str(soup('table')[5].text)=='\n\n\nOn Process.....\n\n\n':
			pass
			
		if str(soup('table')[5].text)!='\n\n\nOn Process.....\n\n\n':
			for i in range(0,9):
			
				if i==0:
					p = 12
					pe = 13
				if i!=0:
					p = p + 7
					pe = pe +7	
							
			        ws.write(row,col,int(soup('table')[7].findAll('font')[p].text))
			        ws.write(row,col+1,int(soup('table')[7].findAll('font')[pe].text))
			        wb.save("Att_Stat_fin.xls")
			        col = col + 3
					
		col = 1			
	        row = row +1
		driver.back()
	wb.save("Att_Stat_fin.xls")
	
	
	


crawl(1,55,"regular",row,col)
row = 58
crawl(31,48,"lateral",row,col)
driver.close()
