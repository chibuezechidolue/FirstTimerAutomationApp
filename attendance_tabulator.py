import pygsheets 
import os
import time

import pygsheets.address


def insert_cell_value(worksheet,cell_label,value,col_num,addr_list):
    current_value=worksheet.get_value(addr=f"{cell_label}{col_num}")
    if current_value== "":
            val_to_update=value
    else:
        val_to_update=f"{current_value} {value}"
    if cell_label!=addr_list[1]:
        worksheet.cell(f"{cell_label}{col_num}").set_horizontal_alignment(pygsheets.custom_types.HorizontalAlignment.CENTER)
    if cell_label==addr_list[3]:
        worksheet.cell(f"{cell_label}{col_num}").set_number_format(pygsheets.custom_types.FormatType.TEXT) 
    worksheet.update_value(addr=f"{cell_label}{col_num}", val=val_to_update, parse=None)
     

def tabulate_result(first_timer_info:list,sheet_name,client,date,addr_list:list):
    """To transfer and tabulate the first_timer_info to an online google sheet"""
    # client.create('First Timers Attendance')
    # client.drive.delete('First Timers Attendance')
    spreadsht = client.open(sheet_name) 
    # spreadsht.share('chinwechidolue@gmail.com', role='writer', type='user', emailMessage='Here is the spreadsheet we talked about!')
    
    worksht = spreadsht.worksheet("title", "Sheet1") 

    adr=pygsheets.address.Address(value=addr_list[0], allow_non_single=True)
    adr_index=adr.index[1]
    col=worksht.get_col(col=adr_index)  # To get the list of values in Col 
    col_num=col.index("")+1
    title=["TTP FIRST TIMER REGISTER",date]
    for n in range(len(title)):
        worksht.merge_cells(start=f"{addr_list[0]}{col_num+n}", end=f"{addr_list[-1]}{col_num+n}", merge_type='MERGE_ROWS', grange=None)
        worksht.cell(f"{addr_list[0]}{col_num+n}").value= title[n]
        worksht.cell(f"{addr_list[0]}{col_num+n}").set_horizontal_alignment(pygsheets.custom_types.HorizontalAlignment.CENTER )
        worksht.cell(f"{addr_list[0]}{col_num+n}").set_text_format(attribute="fontSize", value=16)
        worksht.cell(f"{addr_list[0]}{col_num+n}").set_text_format(attribute="bold", value=True)

    col_titles=['S/N','Name','Gender','Phone no']
    col_label=addr_list
    for n in range(len(col_titles)):
        worksht.cell(f"{col_label[n]}{col_num+2}").value= col_titles[n]   
        worksht.cell(f"{col_label[n]}{col_num+2}").set_horizontal_alignment(pygsheets.custom_types.HorizontalAlignment.CENTER )
        worksht.cell(f"{col_label[n]}{col_num+2}").set_text_format(attribute="fontSize", value=12)
        worksht.cell(f"{col_label[n]}{col_num+2}").set_text_format(attribute="bold", value=True)    

    col_num+=3
    for n in range(len(first_timer_info)):
        label=" "
        for info in first_timer_info[n]:
            if len(info)>1 and check_if_phone_no(info):
                label=addr_list[3]
            elif len(info)>1 and not check_if_phone_no(info):
                label=addr_list[1]
            elif len(info)==1:
                label=addr_list[2]
                if info.upper()=='M':
                    info="MALE"
                elif info.upper()=='F':
                    info="FEMALE"

            worksht.cell(f"{addr_list[0]}{col_num+n}").value= n+1
            worksht.cell(f"{addr_list[0]}{col_num+n}").set_horizontal_alignment(pygsheets.custom_types.HorizontalAlignment.CENTER)
            insert_cell_value(worksheet=worksht,cell_label=label,value=info,col_num=col_num+n,addr_list=addr_list)



def check_if_phone_no(value):
    try:
        int(value)
    except:
        return False
    return True


def convert_txt_to_list(txt_file)->list:
    with open(txt_file, 'r') as in_file:
        stripped = [line.strip() for line in in_file]
        lines = [line.split(" ") for line in stripped if line]
        return lines
    

DATE=input("Please fill in the Date for This attendance: ")

address_list=input('Please type in 4 Column Label(e.g. A,B,C,D): ').upper()

ADDRESS_LIST=address_list.split(",")
print('Trying to authorize the client ...')
client = pygsheets.authorize(service_account_file='JSON/FOLDER/first-timers-attendance.json')
print('client authorized')
print(" Now Tabulating... ")
print('Please do not close the app')

output=convert_txt_to_list(txt_file='first_timers_info.txt')

tabulate_result(addr_list=ADDRESS_LIST,first_timer_info=output,sheet_name="First Timers Attendance",client=client,date=DATE)

print('Tabulation done')
time.sleep(5)
print('Finished !!!')
