# -*- coding: utf-8 -*-
import xlrd
import xlwt
import requests
import tldextract
import whois
import datetime

def excelStart(xlsxName):
    workbook = xlrd.open_workbook(xlsxName)

    worksheet = workbook.sheet_by_index(0)
    nrows = worksheet.nrows


    wtxl = xlwt.Workbook()
    wtxl.default_style.font.height = 20*11

    wtsh = wtxl.add_sheet('A Test Sheet')

    style = xlwt.XFStyle()

    font = xlwt.Font()
    font.bold = True
    font.height = 320
    style.font = font

    wtsh.col(1).width = 256 * 30
    wtsh.col(2).width = 256 * 20
    wtsh.col(3).width = 256 * 15
    wtsh.col(4).width = 256 * 15
    wtsh.col(5).width = 256 * 15

    wtsh.write(0,1,'Domain', style=style)
    wtsh.write(0,2,'IP', style=style)
    wtsh.write(0,3,'Code', style=style)
    wtsh.write(0,4,'Status', style=style)
    wtsh.write(0,5,'Mark', style=style)

    row_ip = []
    row_val = []
    row_status = []
    domain = []
    row_result = []
    x = 0
    for row_all in range(nrows):
        row_ip.append(worksheet.cell_value(row_all, 0))
        row_val.append(worksheet.cell_value(row_all, 2))
        row_status.append(worksheet.cell_value(row_all, 7))


    for i in xrange(0,len(row_val)):
        ext = tldextract.extract(row_val[i])
        domain.append(ext.domain + "." + ext.suffix)

    domain = list(set(domain))

    for y in xrange(0, len(domain)):
        count = 0
        domainUrl = str(domain[y])
        queryWhois = ""
        ExpirationDate = "The End"
        RegistrantName = "not Exist"
        for i in xrange(0,len(row_val)):
            if row_val[i].find(domainUrl) > 0:
                count = count + 1
                x = x + 1
                try:
                    request = requests.get(row_val[i])
                    queryWhois = whois.whois(domainUrl).query()
                    ExpirationDate = whois.Parser(domainUrl, queryWhois[1]).parse()['ExpirationDate']
                    try:
                        RegistrantName = whois.Parser(domainUrl, queryWhois[1]).parse()['RegistrantName']
                    except Exception as e:
                        pass

                    if request.status_code == 200:
                        print(row_ip[i] + "\t\t"+ row_val[i] + " "*10 + "\tFOUND\t\t" + row_status[i])
                        wtsh.write(x,1,row_val[i])
                        wtsh.write(x,2,row_ip[i])
                        wtsh.write(x,4,'FOUND')
                        wtsh.write(x,5,row_status[i])
                    else:
                        print(row_ip[i] + "\t\t" + row_val[i] + " "*10 + "\tNOT(" + str(request.status_code) + ")\t\t" + row_status[i])
                        wtsh.write(x,1,row_val[i])
                        wtsh.write(x,2,row_ip[i])
                        wtsh.write(x,3,str(request.status_code))
                        wtsh.write(x,4,'NOTFOUND')
                        wtsh.write(x,5,row_status[i])

                except requests.exceptions.RequestException as e:
                    print(row_ip[i] + "\t\t" + row_val[i] + " "*10 + "\tERROR\t\t" + row_status[i])
                    wtsh.write(x,1,row_val[i])
                    wtsh.write(x,2,row_ip[i])
                    wtsh.write(x,4,'ERROR')
                    wtsh.write(x,5,row_status[i])
        print("--------------------------------------------------------------------------------------")
        print("Domain: " + domainUrl)
        print("Registrant Name: " + str(RegistrantName))
        print("Expiration Date: " + str(ExpirationDate))
        print("Sub Domain Count: " + str(count))
        print("")
        print("")



    current_time = datetime.datetime.now()
    excelName = str(current_time)[0:10] +'_'+ str(current_time)[11:13] + str(current_time)[14:16] + str(current_time)[17:19]
    excelName = 'Server_' + excelName + '.xls'
    wtxl.save(excelName)