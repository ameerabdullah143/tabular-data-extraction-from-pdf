

import tabula as tb
import pandas as pd
import re
import numpy as np
import sys
import PyPDF2 
import glob, os

path = 'C:/Users/Ameer Abdullah/Desktop/PDF Data extraction'


us_state_to_abbrev = {
    "Canada":"CANADA",
    "Alabama": "AL",
    "Alaska": "AK",
    "Arizona": "AZ",
    "Arkansas": "AR",
    "California": "CA",
    "Colorado": "CO",
    "Connecticut": "CT",
    "Delaware": "DE",
    "Florida": "FL",
    "Georgia": "GA",
    "Hawaii": "HI",
    "Idaho": "ID",
    "Illinois": "IL",
    "Indiana": "IN",
    "Iowa": "IA",
    "Kansas": "KS",
    "Kentucky": "KY",
    "Louisiana": "LA",
    "Maine": "ME",
    "Maryland": "MD",
    "Massachusetts": "MA",
    "Michigan": "MI",
    "Minnesota": "MN",
    "Mississippi": "MS",
    "Missouri": "MO",
    "Montana": "MT",
    "Nebraska": "NE",
    "Nevada": "NV",
    "New Hampshire": "NH",
    "New Jersey": "NJ",
    "New Mexico": "NM",
    "New York": "NY",
    "North Carolina": "NC",
    "North Dakota": "ND",
    "Ohio": "OH",
    "Oklahoma": "OK",
    "Oregon": "OR",
    "Pennsylvania": "PA",
    "Rhode Island": "RI",
    "South Carolina": "SC",
    "South Dakota": "SD",
    "Tennessee": "TN",
    "Texas": "TX",
    "Utah": "UT",
    "Vermont": "VT",
    "Virginia": "VA",
    "Washington": "WA",
    "West Virginia": "WV",
    "Wisconsin": "WI",
    "Wyoming": "WY",
    "District of Columbia": "DC",
    "American Samoa": "AS",
    "Guam": "GU",
    "Northern Mariana Islands": "MP",
    "Puerto Rico": "PR",
    "United States Minor Outlying Islands": "UM",
    "U.S. Virgin Islands": "VI",
    
}
states = {v: k for k, v in us_state_to_abbrev.items()}

def getfileData(df):
    table = []
        
    for page in df:
        page = page.replace(np.nan, '', regex=True)
        for i in range(page.shape[0]):
            if bool(re.match("[A-Z][A-Z]-[0-9]*", page.iloc[i,0])):
                page = page.iloc[i:,:]
                break         
    
        for i in range((page.shape[0])):
            column = []
            if bool(re.match("[A-Z][A-Z]-[0-9]*", page.iloc[i,0])):
                column.append(page.iloc[i,0]) #c1
                column.append(page.iloc[i,1]) #c2
                c3 = (page.iloc[i,2])
                c4 = (page.iloc[i,3])
                c5 = ''
     
                for j in range(i+1,i+8):
                    
                        
                    if page.iloc[j,0]=='':
                        c3 = c3 + ' ' + (page.iloc[j,2])
                        c4 = c4 + ' ' + (page.iloc[j,3])
                        continue
                    elif not (bool(re.match("[A-Z][A-Z]-[0-9]*", page.iloc[j,0]))):
                        if c5!='':
                            c5=c5+', '
                        c5 = c5 + page.iloc[j,0]+page.iloc[j,1]+page.iloc[j,2]+page.iloc[j,3]
                    
                    if j+1 == page.shape[0] or (bool(re.match("[A-Z][A-Z]-[0-9]*", page.iloc[j,0]))):
                        break
                    
                    
                        
                        
                column.append(c3)
                column.append(c4)
                column.append(c5)
                
                table.append(column)
                

    return pd.DataFrame(table)

def findStartPage(file):

    # open the pdf file
    #object = PyPDF2.PdfFileReader(r"LI_REGISTER20211202.pdf")
    object = PyPDF2.PdfFileReader(file)
    
    # get number of pages
    NumPages = object.getNumPages()
    
    # define keyterms
    String = "NON-FITNESS"
    
    # extract text and do the search
    for i in range(0, NumPages):
        PageObj = object.getPage(i)
        Text = PageObj.extractText()
        ResSearch = re.search(String, Text)
        if ResSearch != None:
            return i+1,NumPages
    
            

def get_state_names(data):
    states_list = []
    address = data['APPLICANT'].reset_index(drop=True)
    for i in range(len(address)):
        st = address[i].split(',')[1]
        for k in states.keys():
            if k in st:
                states_list.append(states[k])
                break
        
    data['States'] = states_list
    return data




#print('Write a complete path where your files are located: ')
#path = input()

path = 'C:/Users/Ameer Abdullah/Desktop/PDF Data extraction'
os.chdir(path)
files = glob.glob("*.pdf")

if len(files)==0:
    print('Please write a valid path where files are located!')
    sys.exit()



columns = ['NUMBER','FILED','APPLICANT','REPRESENTATIVE','Business']



    

table = pd.DataFrame()


for file in files:
    start,end = findStartPage(file)
    pages = str(start)+'-'+str(end)
    df= tb.read_pdf(file, pages = pages, area = (20, 20, 750, 950), columns = [80, 150, 330], pandas_options={'header': None}, stream=True)
    
    
    t = getfileData(df)
#    t.columns = columns

    table = table.append(t)
#    table.loc[len(table)] = t

    
output = 'Raw Data.xlsx'

#table = pd.DataFrame(table)
table.columns = columns
#table = table.reset_index()
table.to_excel(output,index=False)



data_with_states = get_state_names(table)


#writer = pd.ExcelWriter('State_Wise_Data.xlsx',engine='xlsxwriter')   


with pd.ExcelWriter('State_wise_data.xlsx', engine='xlsxwriter') as writer:
    for i in pd.unique(data_with_states['States']):
        df = data_with_states[data_with_states['States']==i]
        df = df.drop('States',axis=1).sort_values('Business',axis=0)
        # workbook=writer.book
        # worksheet=workbook.add_worksheet(i)
        # writer.sheets[i] = worksheet
        df.to_excel(writer,sheet_name=i,index=False)   


 