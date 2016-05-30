import re
import shutil
import os
from os import listdir
import xlrd
import xlwt
#open file
#f = open('/samba/allaccess/1.txt', 'r')
#read file
#str = f.read()
#counter


listFF = listdir('/samba/allaccess/SRV/')
#create excel workbook with write permissions (xlwt module)
wb = xlwt.Workbook()
#create sheet IP LIST with cell overwrite rights
ws = wb.add_sheet('IP LIST', cell_overwrite_ok=True)
#create counter
i = 0

def search():
    global i
    #search OS
    OSname = re.search(r"Operating System:.*?([\S\s].*)\s", str).group(1)
    #search SP
    SPname = re.search(r"Service Pack:.*?([\S\s].*)\s", str).group(1)
    #combine OS + " " + SP like "Windows XP Service Pack 1"
    OS_SP = OSname + " " + SPname
    #search domain
    domain = re.search(r"Name:.*?([\S\s].*)\s", str).group(1)
    #search network name
    netbios = re.search(r"The network name.*:.*?([\S\s].*)\s", str).group(1)
    #search IP address+netmask (test purpose only, not used atm)
    ipaddress_mask = re.search(r"IP-address:.*?([\S\s].*)\s", str).group(1)
    #search only IP address
    ipaddress = re.search(r"IP-address:.*?((?:[0-9]{1,3}\.){3}[0-9]{1,3})\/", str).group(1)
    #combine network name + " " + IP address
    netbios_ip = netbios + " " + ipaddress
    #search Windows Installer software
    software_list1 = re.search(r"installation date([\S\s]*)Installed software", str).group(1)
    #search Registry software
    software_list2 = re.search(r"Title version([\S\s]*)Keys installed software", str).group(1)
    #combine 2 software lists into 1
    software_list = software_list1 + " " + software_list2
    #search printer list
    printers = re.search(r"name Port([\S\s]*)Regional settings", str).group(1)
    #search server manufacturer + title (HP,Dell,Vmware etc)
    server_man = re.search(r"Manufacturer: ([\S\s]*)Serial .umber", str).group(1)
    #delete Title:, Name: from manufacturer
    try:
      server_man = server_man.replace('Title:','')
    except:
      pass
    try:
      server_man = server_man.replace('Name:','')
    except:
      pass
    #write information to Excel shells
    ws.write(i, 3, OS_SP)
    ws.write(i, 2, netbios_ip)
    ws.write(i, 4, software_list)
    ws.write(i, 5, domain)
    ws.write(i, 8, printers)
    ws.write(i, 1, server_man)	
    print('info saved')
    #counter+1 to next excel line
    i += 1

    
for file in listFF:
   f = open('/samba/allaccess/SRV/{0}'.format(file), 'r+')
   str = f.read()
   print(i, file)
   search()
   print('ready')
   
wb.save('/samba/allaccess/SRV/SRV.xls')
print('wb saved')
