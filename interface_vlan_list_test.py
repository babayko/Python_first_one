import re
import shutil
import os
from os import listdir
import xlrd
import xlwt

listFF = listdir('C:/python/configs/')
#create excel workbook with write permissions (xlwt module)
wb = xlwt.Workbook()
#create sheet IP LIST with cell overwrite rights
ws = wb.add_sheet('IP LIST', cell_overwrite_ok=True)
#writing first row
ws.col(0).width = 256 * 50
ws.col(1).width = 256 * 27
ws.col(2).width = 256 * 30
ws.col(3).width = 256 * 50
ws.col(4).width = 256 * 36
ws.col(5).width = 256 * 20
ws.col(6).width = 256 * 20
ws.col(7).width = 256 * 20

ws.write(0, 0, 'FileName')
ws.write(0, 1, 'Hostname')
ws.write(0, 2, 'Interface Number')
ws.write(0, 3, 'Description/Nameif')
ws.write(0, 4, 'IP address')
ws.write(0, 5, 'cdp')
ws.write(0, 6, 'static route')
ws.write(0, 7, 'summary(testing purpose)')
#create counter
i = 1

def search():
    #w/o this thing counter do not work inside function
    global i
    #trying to find hostname of the device
    hostname = re.findall(r"\nhostname.([\S\s].*)\n", some_str)
    #if some1 is found
    if hostname != []:
        #remove_duplicates
        hostname_no_duplic = list(set(hostname))
        hostname_str = ' '.join(hostname_no_duplic)
        #write to the excel
        ws.write(i, 1, hostname_str)
    #trying to find all interfaces with IP addresses
    vlanif = re.findall(r"\ninterface ((?:Loopback.*|Tunnel.*|GigabitEthernet.*|Vlan.*|Fast.*\n|Serial.*)[^\!]*ip address (?:[0-9]{1,3}\.){3}[0-9]{1,3} (?:[0-9]{1,3}\.){3}[0-9]{1,3}\s)?", some_str)
    #removing duplicates
    vlanif_no_duplic = list(set(vlanif))
    #sorting interfaces alphabetically
    vlanif_no_duplic.sort()
    #if some interfaces were found
    if vlanif_no_duplic != []:
        #for every interface run new searches
        for item in vlanif_no_duplic:
            #combo-wombo list to string
            item_str = ''.join(item)
            #trying to find interface type and name 
            interface = re.findall(r"(Loopback\S*|Tunnel\S*|GigabitEthernet\S*|Vlan\S*|Fast\S*|Serial\S*)\s", item_str)
            #remove duplicates
            interface_no_duplic = list(set(interface))
            #if something is found then convert to string and write to excel
            if interface_no_duplic != []:            	
                interface_str = ''.join(interface_no_duplic)
                ws.write(i, 2, interface_str)
            #trying to find Cisco ASA interface description (nameif)
            interface_asa = re.findall(r"([\S\s]*)\nnameif", item_str)
            #if something is found then convert to string and write to excel
            if interface_asa != []:            	
                interface_asa_str = ''.join(interface_asa)                
                ws.write(i, 2, interface_asa_str)
            #trying to find Cisco switch/router/industrial switch interface description
            descr = re.findall(r"description ([\S\s]*?)\n", item_str)
            #if something is found then convert to string and write to excel
            if descr != []:
                descr_str = ''.join(descr)
                ws.write(i, 3, descr_str)
            #ne pomnu zachem eto pisal, nado budet potom razobratsya)
            nameif = re.findall(r"nameif ([\S\s]*?)\n", item_str)
            #if something is found then convert to string and write to excel
            if nameif != []:
                nameif_str = ''.join(nameif)
                ws.write(i, 3, nameif_str)
            #trying to find IP address terminated on the interface
            ip = re.findall(r"ip address ((?:[0-9]{1,3}\.){3}[0-9]{1,3} (?:[0-9]{1,3}\.){3}[0-9]{1,3})\s?$", item_str)
            #if something is found then convert to string and write to excel
            if ip != []:
                ip_str = ''.join(ip)
                ws.write(i, 4, ip_str)
            #trying to find secondary IP address terminated on the interface
            ip_sec = re.findall(r"ip address ((?:[0-9]{1,3}\.){3}[0-9]{1,3} (?:[0-9]{1,3}\.){3}[0-9]{1,3}) secondary", item_str)
            if ip_sec != []:
                ip_sec_str = ''.join(ip_sec)
                ip_secondary_str = ip_sec_str + ' secondary'
                #counter +1
                i += 1
                ws.write(i, 4, ip_secondary_str)
            #debug thing, write raw interface string found in another excel column, for debug purposes
            ws.write(i, 7, item_str)
            #counter +1
            i += 1

    print('info saved')
    #counter+1 to next excel line
    i += 1

    
for file in listFF:
   f = open('C:/python/configs//{0}'.format(file), 'r+')
   str
   type(str)
   some_str = f.read()
   print(i, file)
   #write filename in first column
   ws.write(i, 0, file)
   search()
   print('ready')
   
wb.save('C:/python/outputdir/interface_vlan_list_test.xls')
print('wb saved')
