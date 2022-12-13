from openpyxl import load_workbook
wb_search = load_workbook('to_data.xlsx')
sheet = wb_search.active
data = [['GX-0016', None, '24.10.2022', 'jt5148210894566', 0.565, 3.5, '0', 'ОПЛАЧЕНО'], ['GX-0016', 17, '24.10.2022', 'jt5147752127949', '0', '0', '0', 'ОПЛАЧЕНО'], ['GX-0016', 62, '24.10.2022', 'jt5147773335024', 0.385, 3.5, '0', 282]]

# print(len(data))
sheet[3][1].value = '222'
# wb_search.save('to_data.xlsx')
a = 0


# for i in range(2, data+1):







for i in range(2, len(data)+1):    

    if a == len(data)+1:
        d = data[a]
        sheet[i][0].value = d[1]
        # print(sheet[i][3].value)
        sheet[i][1].value = d[2]
        sheet[i][2].value = d[3]
        sheet[i][3].value = d[4]
        sheet[i][4].value = d[5]
        sheet[i][5].value = d[6]
        sheet[i][6].value = d[7]
        break
    else:
        print(i)
        d = data[a]
        sheet[i][0].value = d[1]
        # print(sheet[i][3].value)
        sheet[i][1].value = d[2]
        sheet[i][2].value = d[3]
        sheet[i][3].value = d[4]
        sheet[i][4].value = d[5]
        sheet[i][5].value = d[6]
        sheet[i][6].value = d[7]
        a+=1
wb_search.save('to_data.xlsx')