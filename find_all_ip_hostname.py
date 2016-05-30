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


listFF = listdir('C:/python/configs/')
#create excel workbook with write permissions (xlwt module)
wb = xlwt.Workbook()
#create sheet IP LIST with cell overwrite rights
ws = wb.add_sheet('IP LIST', cell_overwrite_ok=True)
#writing first row
ws.col(0).width = 256 * 50
ws.col(1).width = 256 * 27
ws.col(2).width = 256 * 60
ws.col(3).width = 256 * 20
ws.col(4).width = 256 * 40
ws.col(5).width = 256 * 20
ws.col(6).width = 256 * 20
ws.col(7).width = 256 * 20

ws.write(0, 0, 'FileName')
ws.write(0, 1, 'Hostname')
ws.write(0, 2, 'IP address list')
ws.write(0, 3, 'IP address MOXA')
ws.write(0, 4, 'Firmware')
ws.write(0, 5, 'cdp')
ws.write(0, 6, 'static route')
#create counter
i = 1

def search():
    global i
#    #search OS
#    OSname = re.search(r"Operating System:.*?([\S\s].*)\s", str).group(1)
#    #search SP
#    SPname = re.search(r"Service Pack:.*?([\S\s].*)\s", str).group(1)
#    #combine OS + " " + SP like "Windows XP Service Pack 1"
#    OS_SP = OSname + " " + SPname
#    #search domain
#    domain = re.search(r"Name:.*?([\S\s].*)\s", str).group(1)
#    #search network name
#    netbios = re.search(r"The network name.*:.*?([\S\s].*)\s", str).group(1)
    #search_hostname
    hostname = re.findall(r"hostname.([\S\s].*)\n", str)
    if hostname != []:
        #remove_duplicates
        hostname_no_duplic = list(set(hostname))
        hostname_str = ' '.join(hostname_no_duplic)
        ws.write(i, 1, hostname_str)
    #search IP address Cisco interface vlan
    vlanif = re.findall(r"\n(interface Vlan[\S\s].*[\s]*.*[\s]*.*[\s]*.*[\s]*.*[\s]*.*)!", str)
    if vlanif != []:
    #    vlanif_str = list(set(vlanif))
        #combo-wombo list to string
        vlanif_str = ' '.join(vlanif)
        #write to excel
        ws.write(i, 2, vlanif_str)
    #search IP address MOXA
    ipaddress_moxa = re.findall(r"IPAddress.*?((?:[0-9]{1,3}\.){3}[0-9]{1,3})\s", str)
    if ipaddress_moxa != []:
        ipmoxa_str = ' '.join(ipaddress_moxa)
        ws.write(i, 3, ipmoxa_str)
    #search software_version
    soft = re.findall(r"flash:\/{1,2}([\S\s].*).bin", str)
    if soft != []:
        soft_str = ' '.join(soft)
        ws.write(i, 4, soft_str)
    #search cdp_nei
    cdp = re.findall(r"Platform  Port ID([\S\s]*?)#", str)
    if cdp != []:
        cdp_str = ' '.join(cdp)
        ws.write(i, 5, cdp_str)
    #search loopback
    #lpb = re.findall(r"(Loopback.*)\n*.*\n*.ip.*?((?:[0-9]{1,3}\.){3}[0-9]{1,3})", str)
    #if lpb != []:
    #    lpb_str = ' '.join(lpb)
    #    ws.write(i, 6, lpb_str)
   

#    #combine network name + " " + IP address
#    netbios_ip = netbios + " " + ipaddress
#    #search Windows Installer software
#    software_list1 = re.search(r"installation date([\S\s]*)Installed software", str).group(1)
#    #search Registry software
#    software_list2 = re.search(r"Title version([\S\s]*)Keys installed software", str).group(1)
#    #combine 2 software lists into 1
#    software_list = software_list1 + " " + software_list2
#    #search printer list
#    printers = re.search(r"name Port([\S\s]*)Regional settings", str).group(1)
#    #search server manufacturer + title (HP,Dell,Vmware etc)
#    server_man = re.search(r"Manufacturer: ([\S\s]*)Serial .umber", str).group(1)
#    #delete Title:, Name: from manufacturer
#    try:
#      server_man = server_man.replace('Title:','')
#    except:
#      pass
#    try:
#      server_man = server_man.replace('Name:','')
#    except:
#      pass
    #write information to Excel shells
#    ws.write(i, 3, OS_SP)
#    ws.write(i, 2, netbios_ip)
#    ws.write(i, 4, software_list)
#    ws.write(i, 5, domain)
#    ws.write(i, 8, printers)
#    ws.write(i, 1, server_man)	
    print('info saved')
    #counter+1 to next excel line
    i += 1

    
for file in listFF:
   f = open('C:/python/configs//{0}'.format(file), 'r+')
   str = f.read()
   print(i, file)
   #write filename in first column
   ws.write(i, 0, file)
   search()
   print('ready')
   
wb.save('C:/python/outputdir/IP_hostname_result.xls')
print('wb saved')
