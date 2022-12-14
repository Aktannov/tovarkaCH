from openpyxl import load_workbook
wb_search = load_workbook('your.xlsx')
sheet = wb_search.active
data = [['GX-0016', None, '24.10.2022', 'jt5148210894566', 0.565, 3.5, '0', 'ОПЛАЧЕНО'], ['GX-0016', 17, '24.10.2022', 'jt5147752127949', '0', '0', '0', 'ОПЛАЧЕНО'], ['GX-0016', 62, '24.10.2022', 'jt5147773335024', 0.385, 3.5, '0', 282]]

# print(len(data))
sheet[3][1].value = '222'
# wb_search.save('to_data.xlsx')
a = 0


# for i in range(2, data+1):







for i in range(2, len(data)+2):    

    if a == len(data)+1:
        d = data[a]
        sheet[i][0].value = d[0]
        sheet[i][1].value = d[2]
        sheet[i][2].value = d[3]
        sheet[i][3].value = d[4]
        sheet[i][4].value = d[5]
        sheet[i][5].value = int(d[4]) * int(d[5])
        sheet[i][6].value = d[7]
        break
    else:
        d = data[a]
        sheet[i][0].value = d[0]
        sheet[i][1].value = d[2]
        sheet[i][2].value = d[3]    
        sheet[i][3].value = d[4]
        sheet[i][4].value = d[5]

        if float(d[4]) != 0.0 or float(d[5]) != 0.0:
            # print(d[4], d[5])
            sum = (d[4] * d[5])
            sheet[i][5].value = round(sum, 2)
        else:
            sheet[i][5].value = 0
            
            
        sheet[i][6].value = d[7]
        a+=1
wb_search.save('your.xlsx')