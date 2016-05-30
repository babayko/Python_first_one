import re
import shutil
import os
from os import listdir
import xlrd
import xlwt


listFF = listdir('C:/python/ICS/')
#create excel workbook with write permissions (xlwt module)
wb = xlwt.Workbook()
#create sheet IP LIST with cell overwrite rights
ws = wb.add_sheet('Script_result', cell_overwrite_ok=True)
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
ws.write(0, 0, 'Name', style=style)
ws.write(0, 1, 'Subtheme name', style=style)
ws.write(0, 2, 'Requirement', style=style)
ws.write(0, 3, 'Supplemental Guidance', style=style)
ws.write(0, 4, 'Requirement Enhancements', style=style)
ws.write(0, 5, 'Referencies', style=style)
#not used
#ws.write(0, 6, 'Mask', style=style)
#create counter
i = 1

def search():
    global i
    #search OS
    all_refs = re.findall(r"(\n\d\.\d{1,2} [\S\s]*?)\nZXCCXZ", some_str)
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
        find_all_themes = re.findall(r"(\n\d\.\d{1,2}\.\d{1,2} [\S\s]*?)References\n.*", item_str)
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
            print(reqs_str)
          supp = re.findall(r"Supplemental Guidance\n([\S\s]*?)\n\d\.\d{1,2}\.\d{1,2}\.\d{1,2} ", item_2_str)
          if supp!= []:
            supp_str = ''.join(supp)
            ws.write(i, 3, supp_str)
            print(reqs_str)
          i += 1

            #theme_name = re.findall(r"\n(\d\.\d{1,2} .*)\n", item_str)
            #if theme_name!= []:
            #    theme_name_str = ''.join(theme_name)
            #    ws.write(i, 0, theme_name_str)
            #    i += 1
            #subtheme_name = re.findall(r"\n\d\.\d{1,2}\.\d{1,2} ([\S\s]*?)\n", item_str)
            #if subtheme_name!= []:
            #    subtheme_name_str = ''.join(subtheme_name)
            #    ws.write(i, 1, subtheme_name_str)

    print('info saved')

    
for file in listFF:
   f = open('C:/python/ICS/{0}'.format(file), 'r+')
   str
   type(str)
   some_str = f.read()
   print(i, file)
   search()
   print('ready')


wb.save('C:/python/outputdir/ICS.xls')
print('wb saved')
