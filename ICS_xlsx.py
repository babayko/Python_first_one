import re
import shutil
import os
from os import listdir
import xlrd
import xlsxwriter


listFF = listdir('C:/python/ICS/')
#create excel workbook with write permissions (xlsxwriter module)
#wb = xlwt.Workbook()
workbook = xlsxwriter.Workbook('C:/python/outputdir/demo.xlsx')
#writing first row sizes
ws = workbook.add_worksheet()
ws.set_column('A:A', 50)
ws.set_column('B:B', 30)
ws.set_column('C:C', 50)
ws.set_column('D:D', 50)
ws.set_column('F:F', 50)
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})
#top row settings (bold font)
ws.write(0, 0, 'Name', bold)
ws.write(0, 1, 'Subtheme name', bold)
ws.write(0, 2, 'Requirement', bold)
ws.write(0, 3, 'Supplemental Guidance', bold)
ws.write(0, 4, 'Requirement Enhancements', bold)
ws.write(0, 5, 'Referencies', bold)
#not used
#ws.write(0, 6, 'Mask', style=style)
#create counter
i = 1

def search():
    global i
    #search OS
    all_refs = re.findall(r"(\n\d\.\d{1,2} [\S\s]*?\nZXC)CXZ", some_str)
    #vlanif_no_duplic = list(set(vlanif))
    if all_refs != []:
        #combo-wombo list to string   
        #vlanif_str = ' '.join(vlanif)
      for item in all_refs:
        item_str = ''.join(item)
        theme_name = re.findall(r"\n(\d\.\d{1,2} .*)\n", item_str)
        if theme_name!= []:
          theme_name_str = ''.join(theme_name)
          ws.write(i, 0, theme_name_str)
        #i += 1
        find_all_themes = re.findall(r"(\n\d\.\d{1,2}\.\d{1,2} [\S\s]*?DBDBDB)", item_str)
        for item_2 in find_all_themes:
          item_2_str = ''.join(item_2)
          subtheme_name = re.findall(r"\n\d.\d{1,2}.\d{1,2} (.*)\n", item_2_str)
          if subtheme_name!= []:
            subtheme_name_str = ''.join(subtheme_name)
            ws.write(i, 1, subtheme_name_str)
          reqs = re.findall(r"Requirement\n([\S\s]*?)\n\d\.\d{1,2}\.\d{1,2}\.\d{1,2} ", item_2_str)
          if reqs!= []:
            reqs_str = ''.join(reqs)
            ws.write(i, 2, reqs_str)
          supp = re.findall(r"Supplemental Guidance\n([\S\s]*?)\n\d\.\d{1,2}\.\d{1,2}\.\d{1,2} ", item_2_str)
          if supp!= []:
            supp_str = ''.join(supp)
            ws.write(i, 3, supp_str)
          enh = re.findall(r"Requirement Enhancements\n([\S\s]*?)\n\d\.\d{1,2}\.\d{1,2}\.\d{1,2} ", item_2_str)
          if enh!= []:
            enh_str = ''.join(enh)
            ws.write(i, 4, enh_str)
          refs = re.findall(r"References\n([\S\s]*?)DBDBDB", item_2_str)
          if refs!= []:
            refs_str = ''.join(refs)
            ws.write(i, 5, refs_str)
          i += 1
    print('info saved')

    
for file in listFF:
   f = open('C:/python/ICS/{0}'.format(file), 'r+')
   str
   type(str)
   some_str = f.read()
   print(i, file)
   search()
   print('ready')
workbook.close()

#wb.save('C:/python/outputdir/ICS.xls')
print('wb saved')
