import pandas as pd
import xlrd
import re
import xlsxwriter
import datetime

#Read the table
table = pd.read_excel("/Users/starpark/Desktop/house_keeping.xlsx", skiprows=2, index=False)

#Organize the table
table.drop(columns=['COMPLETE','TYPE','CONDITION','ROOM STATUS','DO NOT DISTURB','OUT OF ORDER','ACCOMMODATION COMMENTS'], inplace = True)
table.rename(columns={'ARRIVAL DATE': 'ARRIVAL', 'ARRIVAL TIME': 'TIME', 'DEPARTURE DATE': 'DEPARTURE', 'FRONTDESK STATUS': 'STATUS'}, inplace=True)
table['SERVICE'] = ""
table['NAME'] = ""
table['ROOM STATUS'] = ""
table = table[['NAME', 'ARRIVAL','TIME','DEPARTURE','ROOM','STATUS','SERVICE','ROOM STATUS']]
time_list = table['TIME'].tolist()
new_time_list = []
for i in time_list:
    if i == "Unknown":
        i = ""
    new_time_list.append(i)
table['TIME'] = new_time_list

#Sort the table
room_list = table['ROOM'].tolist()
new_room_list = []
for i in room_list:
    i = i[:4]
    new_room_list.append(int(re.sub("[^0-9]", '' , i)[:4]))
table['indexNumber'] = new_room_list
table['STATUS'] = pd.Categorical(table['STATUS'],["Turnover","Check-out","Stayover","Check-in","Not Reserved"])
table.sort_values(['STATUS', 'indexNumber'], ascending=True, inplace=True)
table.drop('indexNumber', axis=1,inplace=True)

#Calculate days
start_date = table['ARRIVAL'].tolist()
end_date = table['DEPARTURE'].tolist()
start_date_list = []
end_date_list = []
for date in start_date:
    if str(date) != 'nan':
        start_date_list.append(datetime.datetime.strptime(str(date), "%Y-%m-%d"))
for date in end_date:
    if str(date) != 'nan':
        end_date_list.append(datetime.datetime.strptime(str(date), "%Y-%m-%d"))
days_list = []
for i in range(len(start_date_list)):
        days_list.append(end_date_list[i] - start_date_list[i])
today = datetime.datetime.today().strftime("%Y-%m-%d")
today_date = datetime.datetime.strptime(today, "%Y-%m-%d")
stay_days = []
for i in range(len(start_date_list)):
    stay_days.append(today_date - start_date_list[i])
days = []
for i in range(len(stay_days)):
    days.append(str((stay_days[i]).days))
for i in range(len(days_list)):
    days[i] += '/'
    days[i] += str((days_list[i]).days)
while len(days) != len(room_list):
    days.append('')
table['DEPARTURE'] = days
table.rename(columns={'DEPARTURE': 'DAYS'}, inplace = True)

#Match room to service number
service_one = ['104', '208', '903', '1103', '1208B', '1003', '1204A', '112', '113', '302B', '307', '313', '406', '604A', '604B', '702A', '702B', '706', '802A', '802B', '1104A', '1204B', '1303', '1304A', '1304B', '1403', '1605', '302A']
service_one_half = ['407QS', '701QS', '801QS', '1301Q', '1107', '1207', '1307Q', '1601Q', '1801Q', '1901Q']
service_two = ['205T', '3003T', '2902T', '2501Q', '2302T', '1803T', '1603T', '1502T', '2703T', '3001', '2602T', '2003T', '1310T', '509T', '410T']
service_two_half = ['1804']

status_list = table['STATUS'].tolist()
clean_rooms = 0
service_rooms = 0
clean_rooms += status_list.count('Turnover')
clean_rooms += status_list.count('Check-out')
service_rooms += status_list.count('Stayover')
total_rooms = clean_rooms + service_rooms
sorted_room_list = table['ROOM'].tolist()
service_list = []
for i in range(0, len(sorted_room_list)):
    room_num = re.sub("[^0-9,A-Z]", '' , sorted_room_list[i][:5])
    if room_num in service_one:
        service_list.append('1')
    if room_num in service_one_half:
        service_list.append('1.5')
    if room_num in service_two:
        service_list.append('2')
    if room_num in service_two_half:
        service_list.append('2.5')
table['SERVICE'] = service_list

#Match room to categories
one_bed = ['302A', '304A', '604A', '702A', '802A', '804A', '904A', '1104A', '1304A', '112', '113', '208', '307', '313', '406', '503', '702B', '802B', '903', '1003', '1108A', '1108B', '1208B','1303', '1403', '1406', '1605', '706', '604B', '804B', '1104B', '1304B']
two_bed_1 = ['407QS', '1307Q', '701QS', '801QS', '1301Q', '1311Q', '1601Q', '1801Q', '1901Q']
two_bed_2 = ['2501Q', '205T', '509T', '1310T', '1603T', '2003T', '2302T', '2602T', '2703T', '2902T', '3001', '3003T']
three_bed_1 = ['410T', '1502T', '1803T']
three_bed_2 = ['1604T']
status_list = table['STATUS'].tolist()

clean_rooms = 0
service_rooms = 0
clean_rooms += status_list.count('Turnover')
clean_rooms += status_list.count('Check-out')
service_rooms += status_list.count('Stayover')

sorted_room_list = table['ROOM'].tolist()
clean_num = 0
clean_bed = 0
for i in range(0, clean_rooms):
    room_num = re.sub("[^0-9,A-Z]", '' , sorted_room_list[i][:5])
    if room_num in one_bed:
        clean_num += 1
        clean_bed += 1
    if room_num in two_bed_1:
        clean_num += 1.5
        clean_bed += 2
    if room_num in two_bed_2:
        clean_num += 2
        clean_bed += 2
    if room_num in three_bed_1:
        clean_num += 2
        clean_bed += 3
    if room_num in three_bed_2:
        clean_num += 2.5
        clean_bed += 3

service_num = 0
service_bed = 0
for i in range(clean_rooms, clean_rooms + service_rooms):
    room_num = re.sub("[^0-9,A-Z]", '' , sorted_room_list[i][:5])
    if room_num in one_bed:
        service_num += 1
        service_bed += 1
    if room_num in two_bed_1:
        service_num += 1.5
        service_bed += 2
    if room_num in two_bed_2:
        service_num += 2
        service_bed += 2
    if room_num in three_bed_1:
        service_num += 2
        service_bed += 3
    if room_num in three_bed_2:
        service_num += 2.5
        service_bed += 3

#Export the table
writer = pd.ExcelWriter('/Users/christinesun/Desktop/house_keeping_result.xlsx', engine='xlsxwriter')
table.to_excel(writer, sheet_name='house_keeping',  index=False)



#Adding format
workbook = writer.book
worksheet = writer.sheets['house_keeping']
cell_format = workbook.add_format(
    {   'align': 'center',
        "border": 1,
        "border_color": "#000000",
        "font_name": 'Arial',
        "font_size": 13
    }
)
backgroud_color = workbook.add_format(
    {   
        'align': 'center',
        "border": 1,
        "border_color": "#000000",
        "bg_color": '#CAC9C9',
        "font_name": 'Arial',
        "font_size": 13
    }
)

####Adding the summary####
len_table = len(sorted_room_list)
#Writing clean rooms
worksheet.write(len_table+1, 0, "Clean Rooms", backgroud_color)
worksheet.write(len_table+2, 0, clean_rooms)
#Writing clean number
worksheet.write(len_table+1, 1,  "Clean Number", backgroud_color)
worksheet.write(len_table+2, 1, clean_num)
#Writing clean beds
worksheet.merge_range(len_table+1, 2, len_table+1, 3, "Clean Beds", backgroud_color)
worksheet.merge_range(len_table+2, 2, len_table+2, 3, clean_bed)
#Writing service rooms
worksheet.write(len_table+1, 4, 'Service Rooms', backgroud_color)
worksheet.write(len_table+2, 4,  service_rooms)
#Writing service number
worksheet.write(len_table+1, 5, "Service Number", backgroud_color)
worksheet.write(len_table+2, 5, service_num)
#Writing service beds
worksheet.write(len_table+1, 6, "Service Beds", backgroud_color)
worksheet.write(len_table+2, 6, service_bed)
#Writing notes
worksheet.write(len_table+1, 7, "Notes", backgroud_color)
worksheet.write(len_table+2, 7, ' ')

#Format the table
worksheet.set_column(0, 0, 14,cell_format) 
worksheet.set_column(1, 1, 16,cell_format)
worksheet.set_column(2, 2, None,cell_format) 
worksheet.set_column(3, 3, None,cell_format) 
worksheet.set_column(4, 4, 19,cell_format)
worksheet.set_column(5, 5, 18,cell_format)
worksheet.set_column(6, 6, 14,cell_format)
worksheet.set_column(7, 7, 17,cell_format)
for row in range(1, len_table+1, 2):
    worksheet.set_row(row, None, backgroud_color)
worksheet.set_landscape()
writer.save()



