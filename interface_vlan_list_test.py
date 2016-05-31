import re
import shutil
import os
from os import listdir
import xlrd
import xlwt


WIDTH_CONST = 256
widths = 50, 27, 30, 50, 36, 20, 20, 20
headers = (
    'FileName', 'Hostname', 'Interface Number', 'Description/Nameif', 'IP address',
    'cdp', 'static route', 'summary(testing purpose)')
    

def find_device_hostname(some_str):
    """trying to find hostname of the device"""
    hostname = ''
    # trying to find hostname of the device
    search_result = re.findall(r"\nhostname.([\S\s].*)\n", some_str)
    # if some1 is found
    if search_result:
        #remove_duplicates
        hostname = search_result
        
    return hostname


def find_all_interfaces_with_ip(some_str):
    """find all interfaces with IP addresses"""
    vlanif = re.findall(
        r"\ninterface ((?:Loopback.*|Tunnel.*|GigabitEthernet.*|Vlan.*|Fast.*\n|Serial.*)[^\!]*ip address "
        "(?:[0-9]{1,3}\.){3}[0-9]{1,3} (?:[0-9]{1,3}\.){3}[0-9]{1,3}\s)?", some_str)
        
    # removing duplicates and sorting
    return sorted(list(set(vlanif)))
    
    
def find_interface_type_and_name(some_str):
    """find interface type and name"""
    # trying to find interface type and name 
    interface = re.findall(r"(Loopback\S*|Tunnel\S*|GigabitEthernet\S*|Vlan\S*|Fast\S*|Serial\S*)\s", some_str)
    
    # remove duplicates
    return list(set(interface))
    

listFF = listdir('C:/python/configs/')
# create excel workbook with write permissions (xlwt module)
wb = xlwt.Workbook()
# create sheet IP LIST with cell overwrite rights
ws = wb.add_sheet('IP LIST', cell_overwrite_ok=True)
# set width
for index, width in enumerate(widths):
    ws.col(index).width = WIDTH_CONST * width
# writing first row
for index, header in enumerate(header):
    ws.write(0, index, header)

#create counter
i = 1

def search():
    #w/o this thing counter do not work inside function
    global i
    # find hostname
    hostname = find_device_hostname(some_str)
    if hostname:
        ws.write(i, 1, hostname)
    # trying to find all interfaces with IP addresses
    vlanif = find_all_interfaces_with_ip(some_str)
    # for every interface run new searches
    for item_str in map(''.join, vlanif):
        interface = find_interface_type_and_name(item_str)
        # if something is found then convert to string and write to excel
        if interface:            
            ws.write(i, 2, ''.join(interface))
        # trying to find Cisco ASA interface description (nameif)
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
   some_str = f.read()
   print(i, file)
   #write filename in first column
   ws.write(i, 0, file)
   search()
   print('ready')
   
wb.save('C:/python/outputdir/interface_vlan_list_test.xls')
print('wb saved')
