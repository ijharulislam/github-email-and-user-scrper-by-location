
import urllib2
import json
from bs4 import BeautifulSoup
import xlsxwriter

data = []
for k in range(1,6):
	response = urllib2.urlopen('https://api.github.com/search/users?q=location%3AMelbourne&page='+"%s"%k)

	json_data =json.load(response)
	items =json_data["items"]
	for i in items:
		output={}
		html_url = i["html_url"]
		output["Url"] = html_url
		output["User Name"] = i["login"]
		if html_url:
			get_email = urllib2.urlopen(html_url)
			soup = BeautifulSoup(get_email, "lxml")
			try:
				email = soup.find("li", attrs={"aria-label":"Email"}).find("a")
				
				output["Email"] = email.text
				data.append(output)
				print data
			except:
				output["Email"] = ""
				continue

def write_to_excel(workbook,worksheet,data):
            
            # w = tzwhere.tzwhere()
            bold = workbook.add_format({'bold': True})
            bold_italic = workbook.add_format({'bold': True, 'italic':True})
            border_bold = workbook.add_format({'border':True,'bold':True})
            border_bold_grey = workbook.add_format({'border':True,'bold':True,'bg_color':'#d3d3d3'})
            border = workbook.add_format({'border':True,'bold':True})
            
            #worksheet = workbook.add_worksheet('%s_%s'%(a,j))
            worksheet.set_column('B:D', 22)
            worksheet.set_column('E:F', 33)
            row = 0
            col = 0


            worksheet.write(row,col,'Store List',bold)
            row = row + 1

            row = row + 2

            worksheet.write(row,col,'Sl No',border_bold_grey)
            col = col + 1
            worksheet.write(row,col,'User Name',border_bold_grey)
            col = col + 1
            worksheet.write(row,col,'Email',border_bold_grey)
            col = col + 1
            worksheet.write(row,col,'Url',border_bold_grey)
            
            row = row + 1
            i = 0

            for output in data:
                    
                i = i + 1
                col = 0
                worksheet.write(row, col, i, border)
                col = col + 1
                worksheet.write(row, col, output["User Name"] if output.has_key('User Name') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Email"] if output.has_key('Email') else '',border)
                col = col + 1
                worksheet.write(row, col, output["Url"] if output.has_key('Url') else '',border)
               

                col = col + 1
                row = row + 1

workbook = xlsxwriter.Workbook('github.xlsx')
worksheet = workbook.add_worksheet('Melbourne Programmers')
write_to_excel(workbook,worksheet,data)
workbook.close()
