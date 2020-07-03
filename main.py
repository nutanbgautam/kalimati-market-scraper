import requests
from bs4 import BeautifulSoup
import xlsxwriter
from datetime import timedelta, date

total_no_of_days = 0
total_no_of_days_data = 0


def request_data(URL, data_date):
    BODY = {
        'cdate': data_date,
        'pricetype': 'W'
    }
    r = requests.post(url=URL, data=BODY)
    r.close()
    return r


def filter_request_data(r):
    soup = BeautifulSoup(r.text, 'html.parser')
    trc = soup.find('tr', {'class': 'trc'}) #Heading table row
    row0 = soup.findAll('tr', {'class': 'row0'}) #Data table row 1
    row1 = soup.findAll('tr', {'class': 'row1'}) #Data table row 2
    row0td = [result.findAll('td') for result in row0]
    row1td = [result.findAll('td') for result in row1]
    datas0 = [[td.text for td in tds] for tds in row0td]
    datas1 = [[td.text for td in tds] for tds in row1td]
    return trc, datas0, datas1


def write_to_excel(data_date, heading, datas0, datas1):
    if(datas0 == [] and datas1 == []):
        workbook = xlsxwriter.Workbook(
            data_date.replace("/", "_")+" - NO DATA")
    else:
        global total_no_of_days_data
        workbook = xlsxwriter.Workbook(data_date.replace("/", "_")+".xlsx")
        worksheet = workbook.add_worksheet(data_date.replace("/", " "))
        #Excel Formatting
        bold = workbook.add_format({'bold': True})

        row = 0
        col = 0

        for data_heading in heading:
            worksheet.write(row, col, data_heading.text, bold)
            col += 1
        row += 1

        for i in range(int((len(datas0)+len(datas1))/2)):
            col = 0
            for data in datas1[i]:
                worksheet.write(row, col, data)
                col += 1
            col = 0
            row += 1
            for data in datas0[i]:
                worksheet.write(row, col, data)
                col += 1
            row += 1

        total_no_of_days_data += 1

    workbook.close()


def get_all_dates(sd, ed):
    for n in range(int((ed-sd).days)+1):
        yield sd+timedelta(n)


def main():
    global total_no_of_days, total_no_of_days_data
    URL = "http://kalimatimarket.gov.np/priceinfo/dlypricebulletin"
    start_date = date(2019, 4, 14)
    end_date = date(2020, 4, 12)
    try:
        for dt in get_all_dates(start_date, end_date):
            print("Date Reached = "+str(dt.strftime("%m/%d/%Y")))
            request_date = (str(dt.strftime("%m/%d/%Y")))
            r = request_data(URL, request_date)
            trc, datas0, datas1 = filter_request_data(r)
            write_to_excel(request_date, trc, datas0, datas1)
            total_no_of_days += 1
    except:
        print("Some error occured")

    print("\n"+str(total_no_of_days_data)+" days data collected out of " +
          str(total_no_of_days)+" days")


main()
