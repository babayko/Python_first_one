import re
import shutil
import os
from os import listdir
import xlrd
import xlwt


listFF = listdir('C:/python/cpinfo/')
#create excel workbook with write permissions (xlwt module)
wb = xlwt.Workbook()
#create sheet IP LIST with cell overwrite rights
ws = wb.add_sheet('IP LIST', cell_overwrite_ok=True)
#writing first row sizes
ws.col(0).width = 256 * 50
ws.col(1).width = 256 * 25
ws.col(2).width = 256 * 30
ws.col(3).width = 256 * 23
ws.col(4).width = 256 * 23
ws.col(5).width = 256 * 23

#style for first excel row
style = xlwt.XFStyle()

# font
font = xlwt.Font()
font.bold = True
style.font = font

# borders
borders = xlwt.Borders()
borders.bottom = xlwt.Borders.DASHED
style.borders = borders

#top row settings (bold font+ dashed border)
ws.write(0, 0, 'FileName', style=style)
ws.write(0, 1, 'Policy Name', style=style)
ws.write(0, 2, 'Hostname', style=style)
ws.write(0, 3, 'Interface', style=style)
ws.write(0, 4, 'IP address', style=style)
ws.write(0, 5, 'Mask', style=style)
#create counter
i = 1

#main function
def search():
    #forcing "i" to work inside this functions
    global i
    
    #search Policy name
    policy = re.findall(b"Policy name:\s*(.*)\n", some_str)
    if policy != []:
        #decode bytes to unicode
        policy_str = ''.join([x.decode("utf-8") for x in policy])
        #write policy to excel
        ws.write(i, 1, policy_str)

    #search hostname of checkpoint
    hostname = re.findall(b"/opt/CPinfo-10/bin/(.*).mod.cpi\n", some_str)
    #if something is found, then do 
    if hostname !=[]:
        #clear duplicates from list
        hostname_final = set(list(hostname))
        #encode bytes->utf8 string
        hostname_final_str = ''.join([x.decode("utf-8") for x in hostname_final])
        #write excel
        ws.write(i, 2, hostname_final_str)

    #search strings with interfaces and addresses
    getifs = re.findall(b"localhost (.* (?:[0-9]{1,3}\.){3}[0-9]{1,3} (?:[0-9]{1,3}\.){3}[0-9]{1,3})\n", some_str)
    #if something is found, then do
    if getifs !=[]:
        #decode bytes to unicode
        getifs_str = ''.join([x.decode("utf-8") for x in getifs])
        #search ifname+ip+mask w/o all other trash
        eth_ip = re.findall(r"(.{1,15} (?:[0-9]{1,3}\.){3}[0-9]{1,3} (?:[0-9]{1,3}\.){3}[0-9]{1,3})", getifs_str)    
        #if something is found, do
        if eth_ip != []:
            #for every item in the list (iface+ip+mask) split items and write to different excel columns
            for item in eth_ip: 
                #search ifname
                eth_name = re.findall(r"(\S*).(?:[0-9]{1,3}\.){3}[0-9]{1,3} (?:[0-9]{1,3}\.){3}[0-9]{1,3}", item)
                #write interface name to excel
                ws.write(i, 3, eth_name)
                #search interfaces
                ipadd = re.findall(r"((?:[0-9]{1,3}\.){3}[0-9]{1,3})\s255", item)
                #write ipaddress to excel
                ws.write(i, 4, ipadd)
                #search netmask
                mask = re.findall(r"((?:[255]{1,3}\.){2}[0-9]{1,3}\.[0-9]{1,3})", item)
                #write netmask to excel
                ws.write(i, 5, mask)
                #counter +1 for next excel row
                i += 1
    print('info saved')



#main script    
for file in listFF:
   #open file from listdir cpinfo in read_binary mode
   f = open('C:/python/cpinfo//{0}'.format(file), 'rb')
   #This. Is. Needed. No comments.  
   str
   type(str)
   #str = readed file
   some_str = f.read()
   #print number row, name of a file
   print(i, file)
   #write filename in first column
   ws.write(i, 0, file)
   #starting main function
   search()
   print('ready')
#saving workbook
wb.save('C:/python/outputdir/cpinfo_interface_list.xls')
print('wb saved')
