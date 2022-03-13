import pydnsbl
import ipaddress
from tabulate import tabulate
import xlwt 
from xlwt import Workbook
from ipwhois import IPWhois
#import win32com.client as win32


'''Output Excel Sheet'''

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

'''Input txt File'''

with open('ipList.txt') as f:
    lines = f.read()

'''IP repo check list'''

ip_checker = pydnsbl.DNSBLIpChecker()
ipList=[]
ipList= lines.split()

'''Writing output on the Excel sheet'''

sheet1.write(0,0,"IP ADDRESS")
sheet1.write(0,1,"STATUS")
sheet1.write(0,2,"REPUTATION SCORE")
sheet1.write(0,3,"NAME")
sheet1.write(0,4,"STATE")
sheet1.write(0,5,"COUNTRY")
sheet1.write(0,6,"ADDRESS")
sheet1.write(0,7,"DESCRIPTION")

for i in range (0,(len(ipList))):

    if ipaddress.ip_address(ipList[i]).is_private:
        sheet1.write(i+1,0,ipList[i])
        sheet1.write(i+1,1,"IP IS PRIVATE")

    else:
        if len((str(ip_checker.check(ipList[i])).split())) >3:
            sheet1.write(i+1,0,ipList[i])
            sheet1.write(i+1,1,str(ip_checker.check(ipList[i])).split()[2])
            sheet1.write(i+1,2,str(ip_checker.check(ipList[i])).split()[3][:-1])        
        else:
            sheet1.write(i+1,0,ipList[i])
            sheet1.write(i+1,1,"NOT BLACKLISTED")
            sheet1.write(i+1,2,str(ip_checker.check(ipList[i])).split()[2][:-1])
        IpObj = IPWhois(ipList[i])
        res=IpObj.lookup_whois()
        details = res['nets'][0]
        sheet1.write(i+1,3,details['name'])
        sheet1.write(i+1,4,details['state'])
        sheet1.write(i+1,5,details['country'])
        sheet1.write(i+1,6,details['address'])
        sheet1.write(i+1,7,details['description'])
#sheet1.Columns.AutoFit()
wb.save('IPrepo.xls')

#print(ip_checker.check('8.8.8.8'))