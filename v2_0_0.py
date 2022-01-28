'''
   Copyright 2021 Maksim Trushin  PET-Technology Balashikha

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
'''
#pip install openpyxl xlsxwriter xlrd
import pandas as pd
import glob, fitz, os, openpyxl, re, datetime
import time
import numpy as np

#pip freeze > requirements.txt
#pip install -r requirements.txt

#parser
#------------------------------------------------------------------------------
#functions
def pars(filename):
    '''returns list of words from pdf 
    '''
    t = []
    tmp = fitz.open(filename)      #open pdf file
    for i in range(tmp.page_count):
        tm = tmp.load_page(i)          #loading page
        t += tm.get_text("text").split()  #parsing text into list
    return t


#------------------------------------------------------------------------------
#names of recipients

def recipients_name(s):
    if 'KK' in s or 'Архив' in s: return 'Лаборатория контроля качества'
    elif 'OLD' in s: return 'Отдление лучевой диагностики ООО "Рога и Копыта"'
    else: return 'фелиал ООО «РиК» 

#------------------------------------------------------------------------------
#date from serial

def date(serial):
    return serial[-6:][0:2] + '.' + serial[-6:][2:4] + '.20' + serial[-6:][4:]

#------------------------------------------------------------------------------
#date of damand inwoise

def demand_invoice(serial, flac):
    if 'KK' in flac or 'Архив' in flac: return f'№{serial[0:3]} от {date(serial)}'
    elif 'OLD' in flac: return f'№{serial[0:3]} от {old_date(serial)}'
    else: return ''

def old_date(serial):
    d = datetime.datetime.strptime(date(serial), '%d.%m.%Y')
    weekd = d.weekday()
    if weekd == 0: d -= datetime.timedelta(days = 3)
    elif weekd == 6: d -= datetime.timedelta(days = 2)
    else: d -= datetime.timedelta(days = 1)
    return d.strftime('%d.%m.%Y')



#------------------------------------------------------------------------------
#typer

def tip(s):
    '''returns type of drug based on seres of drug'''
    if 'F1' in s: return 'ФДГ, 18F'
    elif 'F2' in s: return 'ФЭТ, 18F'
    elif 'F3' in s: return 'ПСМА, 18F'
    elif 'F4' in s: return 'ФЭС, 18F'
    elif 'F5' in s: return 'ФЛТ, 18F'
    elif 'F6' in s: return 'ДОПА, 18F'
 
#------------------------------------------------------------------------------
#renamer
def humanFileNames():
    '''This function see in the folder with reports from radiochemestry lab.
    
    Chemists don't like naming reports good. This function do it.
    Usualy the chemistry lab makes four reports after drug production.
    This is a report of synthesis, bulk, filling and bubble point test.
    The function looks through each file and makes perfect names
    'Synthesis Report', 'BPT Report', 'BULK Report', 'ОТЧЁТ ФАСОВКИ'
    
    
    '''
    pdf_files = glob.glob('*.pdf')               #making list of files by *.pdf mask
    lables = ['Synthesis Report', 'BPT Report', 'BULK Report', 'ОТЧЁТ ФАСОВКИ']
    for i in pdf_files:
        pageText = pars(i)                       #list of words from pdf            
        tmp = pageText[0]+' '+pageText[1]        #parse name from pdf                           
        for j in lables:
            if tmp == j:
                os.rename(i, tmp + '.pdf')      #gives human name for file
    return glob.glob('*.pdf')

#------------------------------------------------------------------------------
#parser from synthesis report
def synth(lst):
    '''
    This function takes file name of sinthesis report and 
    returns activity from cyclotron and time of sinthesis in s in list [act, s]
    
    
    '''
    try:
        a = lst.index('MBq')            #index of cyclotron activity
    except ValueError:
        print('В отчете не найдена активность изотопа')
    try:
        m = lst.index('min.')           #index of time of sinthesis       
    except ValueError:
        print('В отчете не найдена продолжительность синтеза')
    
    act, t = lst[a-1], int(lst[m - 1])*60 + int(lst[m + 1])
                            # minutes is before word 'min.', seconds after 
    return [act, t]

#------------------------------------------------------------------------------
#parser from synthesis report
def bulkParser(lst): 
    '''
    This function takes file name of bulk report and 
    returns activity of drug package, volume of drug,time of 
    sertification, activity of reminds and its volume in list [a, v, t, a, v]
    
    '''
    actInd,actLst, volLst, timeLst = [], [], [], []
    Act = [i for i in lst if i.isdigit()] #finding int activities in
    k=0                                         #BULK report  by integers                             
    for i, val in enumerate(lst):
        for j in Act:
            if val  == j: 
                actInd.append(i)               #indeces of activities
                actLst.append(val)             #list of activities from report
                volLst.append(lst[i-1])  #list of volumes from report
                timeLst.append(lst[i-2]) #list of times from report
                break
    tmp = [i  for i in actLst if (int(i)/1000 > 20)][-1] #last big activity before dilution
    #BULK_Activity = tmp #actLst.index(tmp)
    #volume = volLst[actLst.index(tmp)]
    #ostAct = actLst[actLst.index(tmp) + 1]
    #ostVolume = volLst[actLst.index(tmp) + 1]
    return [tmp, volLst[actLst.index(tmp)], timeLst[actLst.index(tmp) + 1],
            actLst[actLst.index(tmp) + 1], volLst[actLst.index(tmp) + 1]]

#------------------------------------------------------------------------------
#finding LKK vials in package list
def vials(lst, maskLKK = 'ARHIV', maskOLD = 'OLD'):
    '''
    list is the list ow words from package pdf page
    checking mask KK* or ARHIV* vials names, activities  and returns list of LKK vials
    ckecking mask DD.D like name of kontrakt vials or with mask "OLD*" and
    returns list of patients vials
    function returns [[LKK vials name, LKK vial activity, LKK vial volume, 
                       LKK vial time] list], [patients[...like LKK] list]]
    
    
    '''
    return [[[i, lst[lst.index(i)-7], lst[lst.index(i)-2], lst[lst.index(i)-3]] 
                 for i in lst if (re.fullmatch(r'[K]{2}[0-9]', i)) or (maskLKK in i)],
            [[i, lst[lst.index(i)-7], lst[lst.index(i)-2], lst[lst.index(i)-3]] 
                 for i in lst if ((re.fullmatch(r'[0-9]{2}[.]{1}[0-9]{1}', i)) 
                                       or (maskOLD in i)) and (i != '11.4')],
            list(set([i for i in lst if 'F1B' in i]))[0]
           ]
#------------------------------------------------------------------------------
#ARHIV in rus
def arhRUS(lst):
    for i, val in enumerate(lst): 
        if 'ARH' in val[0] : 
            try:
                int(val[0][-1])
                lst[i][0] = 'Архив ' + val[0][-1]
            except ValueError:
                lst[i][0] = 'Архив'
    return lst

#------------------------------------------------------------------------------
#ostatok volume check
def ostvol(s):
    return s[1:] if '-' in s else s

def theodorico():
    reportFilesList = humanFileNames()
    #print(reportFilesList)
    cyclotronAct, duration = synth(pars('Synthesis Report.pdf'))
    BULK_Activity, volume, time_of_sert, ostAct, ostVolume = bulkParser(pars('BULK Report.pdf'))
    LKK_vials, OLD_vials, serial = vials(pars('ОТЧЁТ ФАСОВКИ.pdf'))
    LKK_vials = arhRUS(LKK_vials)
    journal = { 'Активность, МБк' : cyclotronAct,
        'РФП или РВ(агрегатное состояние)' : tip(serial),
        '№ серии' : serial,
        'Время синтеза, мин' : str(duration // 60),#.replace('.',','),
        'Суммарная активность, МБк' : BULK_Activity,
        'Объем РФП, мл' : volume.replace('.',','),
        'Время паспортизации': time_of_sert,
        'Номер паспорта' : serial[0:7] if 'BK' in serial else serial[0:6],
      }
    df_journal = pd.DataFrame([journal]) #vials variables
    flacs = {
            '№ фасов-ки': [i[0] for i in LKK_vials + OLD_vials], 
            'Активность фасовки, МБк': [round(float(i[1])) for i in LKK_vials + OLD_vials],
            'Объем фасовки, мл': [i[2].replace('.', ',') for i in LKK_vials + OLD_vials],
            'Кому выдано, наименование получателя' : [recipients_name(i[0]) for i in LKK_vials+OLD_vials],
            '№ и дата требования/ заказ-заявки': [demand_invoice(serial, i[0]) 
                                                  for i in LKK_vials+OLD_vials],
            'Время выдачи': [i[3] for i in LKK_vials + OLD_vials]
    }
    df_flacs = pd.DataFrame(flacs)
 #contamination of DataFrames
    df_final_journal =pd.concat([df_journal, df_flacs, 
                                 pd.DataFrame([{'Активность остатка, МБк':ostAct, 
                                                'Объем остатка РФЛП, Мл': ostvol(ostVolume).replace('.',',')
                                               }
                                              ]
                                             )
                                ],axis=1)

#Deleting NaNs from Dataframe
    df_final_journal.replace(np.NaN, '', inplace=True)
#creating exel report file    
    df_final_journal.to_excel('./report.xlsx', index = False)
    return df_final_journal

#clio functions
#------------------------------------------------------------------------------
#renamer
def chumanFileNames():
    '''This function see in the folder with reports from radiochemestry lab.
    
    Chemists don't like naming reports good. This function do it.
    Usualy the chemistry lab makes four reports after drug production.
    This is a report of synthesis, bulk, filling and bubble point test.
    The function looks through each file and makes perfect names
    'Synthesis Report', 'BPT Report', 'ОТЧЁТ ФАСОВКИ'
    
    
    '''
    pdf_files = glob.glob('*.pdf')               #making list of files by *.pdf mask
    lables = ['Synthesis Report', 'Bubble Point', 'Distribution report']
    for i in pdf_files:
        pageText = pars(i)                       #list of words from pdf            
        tmp = pageText[0]+' '+pageText[1]        #parse name from pdf                           
        for j in lables:
            if tmp == j:
                os.rename(i, tmp + '.pdf')      #gives human name for file
    return glob.glob('*.pdf')
#------------------------------------------------------------------------------
#parser from synthesis report
def cbulkParser(lst): 
    '''
    This function takes file name of bulk report and 
    returns activity of drug package, volume of drug,time of 
    sertification, activity of reminds and its volume in list [a, v, t, a, v]
    
    '''
    actInd,actLst, volLst, timeLst = [], [], [], []
    volLst = str(float(lst[lst.index('(dilution)') - 2].replace(',', '.')
                  ) + float(lst[lst.index('(predil.)') - 2].replace(',', '.')
                           ) + float(lst[lst.index('SYNTHESIS') - 2].replace(',', '.'))
             ).replace('.', ',')
    actLst = lst[(lst.index('ACTIVITY') -3)]
    timeLst = lst[(lst.index('ACTIVITY') -5)]
    return [actLst, volLst,  timeLst]

#------------------------------------------------------------------------------
#finding LKK vials in package list
def cvials(lst, maskLKK = 'ARHIV', maskOLD = 'OLD'):
    '''
    list is the list ow words from package pdf page
    checking mask KK* or ARHIV* vials names, activities  and returns list of LKK vials
    ckecking mask DD.D like name of kontrakt vials or with mask "OLD*" and
    returns list of patients vials
    function returns [[LKK vials name, LKK vial activity, LKK vial volume, 
                       LKK vial time] list], [patients[...like LKK] list]]
    
    
    '''
    return [[[i, lst[lst.index(i)-8], lst[lst.index(i)-20], lst[lst.index(i)-9]] 
                 for i in lst if (re.fullmatch(r'[K]{2}[0-9]', i)) or (maskLKK in i)
            ],
            [[i, lst[lst.index(i)-8], lst[lst.index(i)-20], lst[lst.index(i)-9]] 
                 for i in lst if ((re.fullmatch(r'[0-9]{2}[.]{1}[0-9]{1}', i)) 
                                       or (maskOLD in i))],
            cserial(lst)
           ]
#------------------------------------------------------------------------------
def cserial(lst):   
    for i in lst:
        if re.fullmatch(r'[0-9]{3}[F]{1}[0-9]{1}[B]{1}[0-9]{6}', i) or re.fullmatch(r'[0-9]{3}[F]{1}[0-9]{1}[BK]{1}[0-9]{6}', i) :
            #print(i)
            break
    return i

def clio():
    reportFilesList = chumanFileNames()
    cyclotronAct, duration = synth(pars('Synthesis Report.pdf'))
    BULK_Activity, volume, time_of_sert = cbulkParser(pars('Distribution report.pdf'))
    #print(cyclotronAct, duration//60, BULK_Activity, volume, time_of_sert)
    LKK_vials, OLD_vials, serial = cvials(pars('Distribution report.pdf'))
    #print(LKK_vials, OLD_vials, serial)
    LKK_vials = arhRUS(LKK_vials)
    journal = { 'Активность, МБк' : cyclotronAct,
        'РФП или РВ(агрегатное состояние)' : tip(serial),
        '№ серии' : serial,
        'Время синтеза, мин' : str(duration // 60),#.replace('.',','),
        'Суммарная активность, МБк' : round(BULK_Activity.replace(',','.')),
        'Объем РФП, мл' : volume,#.replace('.',','),
        'Время паспортизации': time_of_sert,
        'Номер паспорта' : serial[0:7] if 'BK' in serial else serial[0:6],
      }
    df_journal = pd.DataFrame([journal]) #vials variables
    flacs = {
            '№ фасов-ки': [i[0] for i in LKK_vials + OLD_vials], 
            'Активность фасовки, МБк': [i[1] for i in LKK_vials + OLD_vials],
            'Объем фасовки, мл': [i[2] for i in LKK_vials + OLD_vials],
            'Кому выдано, наименование получателя' : [recipients_name(i[0]) for i in LKK_vials+OLD_vials],
            '№ и дата требования/ заказ-заявки': [demand_invoice(serial, i[0]) 
                                                  for i in LKK_vials+OLD_vials],
            'Время выдачи': [i[3] for i in LKK_vials + OLD_vials]
    }
    df_flacs = pd.DataFrame(flacs)
 #contamination of DataFrames
    df_final_journal =pd.concat([df_journal, df_flacs, 
                                ],axis=1)

#Deleting NaNs from Dataframe
    df_final_journal.replace(np.NaN, '', inplace=True)
#creating exel report file    
    df_final_journal.to_excel('./report.xlsx', index = False)
    return df_final_journal

def main():
    if len(glob.glob('*.pdf')) == 4: #theodorico reporter
        theodorico()
    elif len(glob.glob('*.pdf')) == 3: #clio reporter
        clio()

start_time = time.time()
a = main()
end_time = time.time()
os.system('cls')
#os.system('clear')
print(f'время выполнения = {end_time - start_time} секунд')
print ('------------------SUCCES!-------------------', '\n\n\nPress Enter to exit')
input()
