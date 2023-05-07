import datetime
from tkinter import *
from tkinter import font as tkFont
from tkinter import messagebox
import time as tm
import pandas as pd
import xlsxwriter
from datetime import date
#from uuid import getnode as get_mac
import sys
import os
today = date.today()
now = datetime.datetime.now().strftime("%B %d, %Y, %H:%M")
d2 = today.strftime("%B %d, %Y")
error_data = []
root = Tk()
var = IntVar()
root.title('Converting data')
root.geometry("300x150")
root.resizable(0, 0)
entry = Entry(root, width=25, borderwidth=1)
entry.grid(row=0, column=1, sticky=W)
label_speed = Label(root, text="Directory:", fg="black", font=("Helvetica", 13), padx=20, pady=5)
label_speed.grid(row=0, column=0)
button_font = tkFont.Font(family='Helvetica', size=10, weight='bold')
msg_font = tkFont.Font(family='Helvetica', size=12, weight='bold')
#------------

def analysis_err(x):
    global e
    global err
    e = ''
    x = x.split('W')

    for i in range(1, len(x)):
        e = e + 'W' + x[i]
        # if x[i]=='03':
    err = err.append([e])

    return x[0]
#-------------
def write_log(d,ca,auditor,now,error_data):

    f = open('log/data_log.txt','a')
    f.write('{0} -- {1} -- {2} -- {3} -- {4}\n'.format('Date: '+d, 'Shift: '+ca,'Auditor: '+auditor, 'Now: '+now,'Error data: '+error_data))
    f.close()

def run():
    global df, err, lsx, may, stt, vh, time, date, width, err, W03, W04, W05, W06, W07, W08, W09, W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W21, W30, W33, W34
    start_time = tm.time()
    lsx = pd.DataFrame([])
    may = pd.DataFrame([])
    stt = pd.DataFrame([])
    vh = pd.DataFrame([])
    time = pd.DataFrame([])
    date = pd.DataFrame([])
    width = pd.DataFrame([])
    err = pd.DataFrame([])
    W03 = pd.DataFrame([])
    W04 = pd.DataFrame([])
    W05 = pd.DataFrame([])
    W06 = pd.DataFrame([])
    W07 = pd.DataFrame([])
    W08 = pd.DataFrame([])
    W09 = pd.DataFrame([])
    W10 = pd.DataFrame([])
    W11 = pd.DataFrame([])
    W12 = pd.DataFrame([])
    W13 = pd.DataFrame([])
    W14 = pd.DataFrame([])
    W15 = pd.DataFrame([])
    W16 = pd.DataFrame([])
    W17 = pd.DataFrame([])
    W18 = pd.DataFrame([])
    W19 = pd.DataFrame([])
    W21 = pd.DataFrame([])
    W30 = pd.DataFrame([])
    W33 = pd.DataFrame([])
    W34 = pd.DataFrame([])
    t = []

    dir = entry.get()
    if dir == '':
        dir = 'INPUT_DATA.xlsx'

    df = pd.read_excel(dir, dtype=str)
    d = df.loc[0][1]

    if int(df.loc[0][5].split(':')[0]) > 18:
        ca = '2'
    else: ca = '1'
    name = 'REPORT_'+d.replace('/','_')+'_CA_'+ca+'.xlsx'
    auditor = df.loc[0][4]
    print('Auditor: '+auditor)
    print('Date: ' + d)
    print('Shift: '+ ca)

    #for i in range(5, len(df.columns)):
     #   if str(df.loc[0][i]) != 'nan':
      #      t.append(str(df.loc[0][i]))
    #a= len(t)

    for i in range(1, len(df)):

        num = df.loc[i][2]
        # print(df.loc[1][9])
        for j in range(5, len(df.loc[0])):
            data = str(df.loc[i][j]).upper()

            dta_lsx = df.loc[i][0]
            dta_may = df.loc[i][1]
            dta_vh = df.loc[i][4]
            dta_time = df.loc[0][j]
            if ((data == 'NAN') or (data == 'STOP')) or (data == ' '):
                continue
            else:
                if data.find('XC') != -1:
                    if j != 5:
                        num = df.loc[i][3]
                        if str(num) == 'nan':
                            error_data.append(df['Số máy'][i])
                        continue
                else:
                    if data.find('X') != -1:
                        num = df.loc[i][3]
                        if str(num) == 'nan':
                            error_data.append(df['Số máy'][i])

                        data = data.replace('X', '')

                    if data.find('W') == -1:
                        v = int(data)
                        if v < 400 or v > 1400:
                            error_data.append(df['Số máy'][i])
                        width = width.append([data])
                        lsx = lsx.append([dta_lsx])
                        may = may.append([dta_may])
                        stt = stt.append([num])
                        vh = vh.append([dta_vh])
                        time = time.append([dta_time])
                        date = date.append([d])
                        err = err.append([''])
                        W03 = W03.append([''])
                        W04 = W04.append([''])
                        W05 = W05.append([''])
                        W06 = W06.append([''])
                        W07 = W07.append([''])
                        W08 = W08.append([''])
                        W09 = W09.append([''])
                        W10 = W10.append([''])
                        W11 = W11.append([''])
                        W12 = W12.append([''])
                        W13 = W13.append([''])
                        W14 = W14.append([''])
                        W15 = W15.append([''])
                        W16 = W16.append([''])
                        W17 = W17.append([''])
                        W18 = W18.append([''])
                        W19 = W19.append([''])
                        W21 = W21.append([''])
                        W30 = W30.append([''])
                        W33 = W33.append([''])
                        W34 = W34.append([''])



                    else:

                        data = analysis_err(data)
                        v = int(data)
                        if v < 400 or v > 1400:
                            error_data.append(df['Số máy'][i])
                        width = width.append([data])
                        lsx = lsx.append([dta_lsx])
                        may = may.append([dta_may])
                        stt = stt.append([num])
                        vh = vh.append([dta_vh])
                        time = time.append([dta_time])
                        date = date.append([d])
                        W03 = W03.append([''])
                        W04 = W04.append([''])
                        W05 = W05.append([''])
                        W06 = W06.append([''])
                        W07 = W07.append([''])
                        W08 = W08.append([''])
                        W09 = W09.append([''])
                        W10 = W10.append([''])
                        W11 = W11.append([''])
                        W12 = W12.append([''])
                        W13 = W13.append([''])
                        W14 = W14.append([''])
                        W15 = W15.append([''])
                        W16 = W16.append([''])
                        W17 = W17.append([''])
                        W18 = W18.append([''])
                        W19 = W19.append([''])
                        W21 = W21.append([''])
                        W30 = W30.append([''])
                        W33 = W33.append([''])
                        W34 = W34.append([''])
    lsx = lsx.reset_index(drop=True)
    may = may.reset_index(drop=True)
    stt = stt.reset_index(drop=True)
    vh = vh.reset_index(drop=True)
    time = time.reset_index(drop=True)
    date = date.reset_index(drop=True)
    width = width.reset_index(drop=True)
    err = err.reset_index(drop=True)
    W03 = W03.reset_index(drop=True)
    W04 = W04.reset_index(drop=True)
    W05 = W05.reset_index(drop=True)
    W06 = W06.reset_index(drop=True)
    W07 = W07.reset_index(drop=True)
    W08 = W08.reset_index(drop=True)
    W09 = W09.reset_index(drop=True)
    W10 = W10.reset_index(drop=True)
    W11 = W11.reset_index(drop=True)
    W12 = W12.reset_index(drop=True)
    W13 = W13.reset_index(drop=True)
    W14 = W14.reset_index(drop=True)
    W15 = W15.reset_index(drop=True)
    W16 = W16.reset_index(drop=True)
    W17 = W17.reset_index(drop=True)
    W18 = W18.reset_index(drop=True)
    W19 = W19.reset_index(drop=True)
    W21 = W21.reset_index(drop=True)
    W30 = W30.reset_index(drop=True)
    W33 = W33.reset_index(drop=True)
    W34 = W34.reset_index(drop=True)

    result = pd.concat([lsx, may, stt, vh, time, date, width, err, W03, W04, W05, W06, W07, W08, W09,
                        W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W21, W30, W33, W34], axis=1)

    result.columns = ['LSX', 'MAY', 'STT', 'VH', 'TIME', 'DATE', 'WIDTH', 'ERR',
                      'W03', 'W04', 'W05', 'W06', 'W07', 'W08', 'W09', 'W10', 'W11',
                      'W12', 'W13', 'W14', 'W15', 'W16', 'W17', 'W18', 'W19', 'W21', 'W30', 'W33', 'W34']
    result['DATE'] = pd.to_datetime(result.DATE, format='%d/%m/%Y', dayfirst=True).dt.date
    result['TIME'] = pd.to_datetime(result['TIME'], format ='%H:%M')
    result['WIDTH'] = pd.to_numeric(result.WIDTH)
    for i in result.index:
        if result.at[i, 'ERR'].find('W') != -1:
            e2 = result.at[i, 'ERR'].split('W')
            for j in range(0, len(e2)):
                loc = 'W' + e2[j]
                result.at[i, loc] = 'X'
    #print(result.to_string())
    writer = pd.ExcelWriter(name, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    result.to_excel(writer, sheet_name='Report', index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets['Report']

    # Add some cell formats.
    format1 = workbook.add_format()
    format2 = workbook.add_format()
    format3 = workbook.add_format()
    format1.set_font_color('red')
    format2.set_bg_color('yellow')
    format3.set_bg_color('blue')
    format_time = workbook.add_format({'num_format': 'h:mm'})
    # Set the column width and format.
    worksheet.set_column('A:D', None, format1)
    worksheet.set_column('H:H', None, format2)
    worksheet.set_column('AD:AD', None, format2)
    worksheet.set_row(result.index.stop+1, None, format3)
    worksheet.set_column('E:E', None, format_time)
    worksheet.set_column('A:A', 11, None)
    worksheet.set_column('F:F',10,None)
    worksheet.freeze_panes(1, 1)

    writer.save()

    stop_time = tm.time() - start_time
    msg = Message(root, text='Done', pady=10, fg="red")
    msg['font'] = msg_font
    msg.grid(row=1, column=1)
    k = ''
    if len(error_data)>0:
        for i in range(0,len(error_data)):
            k = k+' '+error_data[i]
        print('Recheck at ' + k)
        messagebox.showinfo('Wrong typing - Be careful','Kiểm tra lại nhập liệu của máy: '+'\n'+ k)

    print('Done in ' + str(round(stop_time, 2))+ ' seconds')
    error_data.clear()
    write_log(d,ca,auditor,now,k)

#--------
def getMacAddress():
    if sys.platform == 'win32':
        for line in os.popen("ipconfig /all"):
            if line.lstrip().startswith('Physical Address'):
                mac = line.split(':')[1].strip()
                break
    else:
        for line in os.popen("/sbin/ifconfig"):
            if line.find('Ether') > -1:
                mac = line.split()[4]
                break
    return mac
def check():
    mac = getMacAddress()
    #print(mac)
    mac = 'mac address'
    if mac == 'mac address' or mac == 'mac address':
        run()
    else:
        print('Sorry! The App can only run on two computers of QC-WW.')
        msg = Message(root, text='Sorry! The App can only run on two computers of QC-WW.', pady=10, fg="red")
        msg['font'] = tkFont.Font(family='Helvetica', size=12)
        msg.grid(row=1, column=1)

button_run = Button(root, text='Start',bg='#0052cc', fg='#ffffff', padx=20, pady=10, command=lambda: check())
button_run['font'] = button_font
button_run.grid(row=1, column=0)
msg = Message(root, text=d2, padx=5, pady=30, fg="black", font=("Helvetica", 10))
msg.grid(row=2, column=1)

#
root.mainloop()
