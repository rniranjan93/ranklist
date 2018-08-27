from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import pandas as pd
import re
import os
inp0=str(input("Enter the results url in the format: http://14.139.56.15/scheme16/studentresult/index.asp:"))
url=inp0
inp=str(input("Download 'chromedriver.exe' and specify the path location Example:H:\python\chromedriver.exe:"))
driver = webdriver.Chrome(inp)
driver.implicitly_wait(30)
driver.get(url)
soup_level1=BeautifulSoup(driver.page_source,'lxml')
datalist = []
x = 0
inp1=int(input("Enter your class starting roll number:"))
inp2=int(input("Enter your class ending roll number:"))
inp3=str(input("Enter the file name that won't intersect with the preexisting file name Example:resultscivil5:"))
for i in range(inp1,inp2+1):
    pb=driver.find_element_by_name('B2')
    pb.click()
    username_box=driver.find_element_by_name('RollNumber')
    username_box.send_keys(i)
    pb=driver.find_element_by_name('B1')
    pb.click()
    l=[]
    soup_level2=BeautifulSoup(driver.page_source,'lxml')
    try:
        table = soup_level2.find_all('table')[0]
        table = soup_level2.find_all('table')[5]
        table = soup_level2.find_all('table')[6]
        table = soup_level2.find_all('table')[0]
        df = pd.read_html(str(table),header=0)
        l.append(df[0])
        table = soup_level2.find_all('table')[7]
        df = pd.read_html(str(table),header=0)
        l.append(df[0])
        table = soup_level2.find_all('table')[8]
        df = pd.read_html(str(table),header=0)
        l.append(df[0])
        datalist.append(l)
        driver.execute_script("window.history.go(-1)")
    except:
        driver.execute_script("window.history.go(-1)")
datanew=datalist
cn=[]
cn.append(datalist[0][0].iloc[0,0])
cn.append(datalist[0][0].columns[0])
for i in range(4):
	cn.append(datalist[0][2].columns[i])
k=0
while k+1:
	try:
		cn.append(datalist[0][1].iloc[k,1])
		k=k+1
	except:
		k=-1
h=pd.DataFrame()
t=len(datalist)
for k in range(t):
	l=[]
	l.append(datalist[k][0].iloc[0,1])
	l.append(datalist[k][0].columns[1])
	for i in range(4):
		l.append(datalist[k][2].iloc[0,i])
	m=0
	mdd=0
	while m+1:
		try:
			l.append(datalist[k][1].iloc[m,4])
			m=m+1
			mdd=mdd+1
		except:
			m=-1
	z=pd.DataFrame([l])
	h=h.append(z)
h.index=range(1,t+1)
for i in range(t):
    m=h.iloc[i,2]
    n=h.iloc[i,4]
    k=0
    while k+1:
        if '='==m[k]:
            h.iloc[i,2]=float(m[k+1:])
            k=-1
        else :
            k=k+1
    k=0
    while k+1:
        if '='==n[k]:
            h.iloc[i,4]=float(n[k+1:])
            k=-1
        else :
            k=k+1
h.columns=cn
aciv=h.sort_values('SGPI',ascending=False)
bciv=h.sort_values('CGPI',ascending=False)
aciv.index=range(1,t+1)
bciv.index=range(1,t+1)


z=['No.of students failed in a particular subject']
for c in range(5):
	z.append('')
for c in range(6,mdd+6):
	e=0
	for r in range(t):
		if h.iloc[r,c]=='F':
			e=e+1
	z.append(e)
z=pd.DataFrame([z])
z.columns=aciv.columns
aciv=aciv.append(z)
bciv=bciv.append(z)
z2=['No. of students passed in a paricular subject']
for c in range(5):
	z2.append('')
for c in range(6,mdd+6):
	z2.append(t)
aciv.rename(columns={'SGPI':'C'},inplace=True)
aciv.rename(columns={'CGPI':'S'},inplace=True)
aciv.rename(columns={'C':'CGPI'},inplace=True)
aciv.rename(columns={'S':'SGPI'},inplace=True)
m=pd.DataFrame(aciv.iloc[:,2])
n=pd.DataFrame(aciv.iloc[:,4])
aciv.iloc[:,2]=n.iloc[:,0]
aciv.iloc[:,4]=m.iloc[:,0]
z2=pd.DataFrame([z2])
z2.columns=aciv.columns
z2.iloc[0,6:]=z2.iloc[0,6:]-z.iloc[0,6:]
aciv=aciv.append(z2)
bciv=bciv.append(z2)
writer=pd.ExcelWriter(inp3+('.xlsx'))
bciv.to_excel(writer,'by"CGPI"')
aciv.to_excel(writer,'by"SGPI"')
writer.save()


