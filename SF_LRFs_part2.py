#-------------------------------
"""
Created on Tue Jul 18 14:19:16 2017

@author: ttrinhtran
"""
#-------------------------------

from pandas import Series, DataFrame
import numpy as np
import pandas as pd
import pymssql
import xlsxwriter
import win32com.client as win32

states = {
        'AL': ['Alabama','01'],
        'AK': ['Alaska','02'],
        'AZ': ['Arizona','04'],
        'AR': ['Arkansas','05'],
        'CA': ['California','06'],
        'CO': ['Colorado','08'],
        'CT': ['Connecticut','09'],
        'DE': ['Delaware',10],
        'DC': ['District of Columbia',11],
        'FL': ['Florida',12],
        'GA': ['Georgia',13],
        'HI': ['Hawaii',15],
        'ID': ['Idaho',16],
        'IL': ['Illinois',17],
        'IN': ['Indiana',18],
        'IA': ['Iowa',19],
        'KS': ['Kansas',20],
        'KY': ['Kentucky',21],
        'LA': ['Louisiana',22],
        'ME': ['Maine',23],
        'MD': ['Maryland',24],
        'MA': ['Massachusetts',25],
        'MI': ['Michigan',26],
        'MN': ['Minnesota',27],
        'MS': ['Mississippi',28],
        'MO': ['Missouri',29],
        'MT': ['Montana',30],
        'NE': ['Nebraska',31],
        'NV': ['Nevada',32],
        'NH': ['New Hampshire',33],
        'NJ': ['New Jersey',34],
        'NM': ['New Mexico',35],
        'NY': ['New York',36],
        'NC': ['North Carolina',37],
        'ND': ['North Dakota',38],
        'OH': ['Ohio',39],
        'OK': ['Oklahoma',40],
        'OR': ['Oregon',41],
        'PA': ['Pennsylvania',42],
        'RI': ['Rhode Island',44],
        'SC': ['South Carolina',45],
        'SD': ['South Dakota',46],
        'TN': ['Tennessee',47],
        'TX': ['Texas',48],
        'UT': ['Utah',49],
        'VT': ['Vermont',50],
        'VA': ['Virginia',51],
        'WA': ['Washington',53],
        'WV': ['West Virginia',54],
        'WI': ['Wisconsin',55],
        'WY': ['Wyoming',56]
        }

state = input('Your favorite state: ' ).upper()

line = input('Auto or Home: ' ).capitalize()
while True:
    if line == 'Auto' or line == 'Home':
        break
    else:
        line = input('Auto or Home: ' ).capitalize()

dt = input('Effective date (yyyy-mm-dd): ' )
dt1 = dt.replace('-','_')
st_code = states.get(state)[1]


"""Queries to retrieve data from HDB"""

Census_Query= "Select distinct  a.*,c.INTPTLAT10,c.INTPTLON10,b.ZCTA5CE10 \
from  TIGER_2010.dbf.BLKPOPHU a, TIGER_2010.dbf.FACES b, TIGER_2010.dbf.TABBLOCK c \
Where a.STATEFP10 = %s \
and a.STATEFP10 = b.STATEFP10 \
and b.STATEFP10 = c.STATEFP10 \
and a.COUNTYFP10 = b.COUNTYFP10 \
and b.COUNTYFP10 = c.COUNTYFP10 \
and a.BLOCKCE = b.BLOCKCE10 \
and b.BLOCKCE10 = c.BLOCKCE10 \
and a.TRACTCE10 = b.TRACTCE10 \
and b.TRACTCE10 = c.TRACTCE10 \
and ZCTA5CE10 is not Null \
ORDER BY 1,2,3,4" % st_code

#Zone_Query = "select City, Zip_Code, County_Code, County, Population, CC_Auto_Zone as Zone \
#from HIS_USER2.dbo.Zip_Codes \
#where State = '%s'" % state

Zone_Query = "Select First.Zip_Code, First.County_Code, First.City, First.County, First.Zone, \
First.Population, \
(Case when Second.Wrt_Exp <= 0 or Second.Wrt_Exp is null then 0.00001 else Second.Wrt_Exp end) As Wrt_Exp \
from (select City, Zip_Code, County_Code, County, Population, CC_Auto_Zone as Zone \
from HIS_USER2.dbo.Zip_Codes \
where State = '%s') as First \
LEFT OUTER JOIN (select \
a.STATE as st_name,  s.psc_st_cd, \
a.GARAGE_ZIPCODE, \
sum(wx_imis_mo) as Wrt_Exp \
from Auto_Staging.dbo.PIDHST_VEHICLE a, Actuarial_Reports.dbo.imisprem b, \
Actuarial_Reports.dbo.dim_states s \
where s.psc_st_cd = '%s' \
and a.POLICY_POINTER = b.policy_pointer \
and a.state = s.psc_st_nbr \
and b.policy_view_date between a.HIST_BEG_EFF_DATE and a.HIST_END_EFF_DATE \
and a.UNIT = b.unit \
and year(b.month_end)= 2016 \
and b.major_peril = '110' \
group by  a.STATE, s.psc_st_cd, \
a.GARAGE_ZIPCODE) as Second \
ON First.Zip_Code = Second.GARAGE_ZIPCODE \
order by First.Zip_Code" % (state, state)

"""Get Census, Zone, SF LRFs and lat & long for zip codes not in Census"""

con = pymssql.connect(server="xx1")
Census_df_org = pd.read_sql_query(Census_Query, con)
Census_df = Census_df_org
Zone_df = pd.read_sql_query(Zone_Query, con)
LRFs_df = pd.read_csv(r'T:\ACT_PID\Pcm_brm\Misc\0.01 SF_LRFs\TT_Test\LRFs.csv')
LRFs_df.columns[0] = ['GRID_ID']
missing_df = pd.read_csv(r'T:\ACT_PID\Pcm_brm\Misc\0.01 SF_LRFs\TT_Test\missing.csv')
missing_df = missing_df.ix[missing_df['STATEFP10']==int(st_code)]

"""Get SF LRFs for each Census record; Calculate Min, Max, Average for each unique Zip in Census"""

Census_df['INTPTLAT10'] = Census_df['INTPTLAT10'].str[1:]
Census_df['INTPTLON10'] = Census_df['INTPTLON10'].str[1:]
Census_df['INTPTLAT10'] = Census_df['INTPTLAT10'].astype(float)
Census_df['INTPTLON10'] = Census_df['INTPTLON10'].astype(float)
Zone_df.Zip_Code = Zone_df.Zip_Code.astype(int)
#Zone_df.County_Code = Zone_df.County_Code.astype(int) #this causes problem when there're out-of-state zip codes

Census_df['GEO ID'] = (Census_df['INTPTLAT10'] * 1000 // 10)*100000 + Census_df['INTPTLON10'] * 1000 // (10)

Census_df['GEO ID'] = Census_df['GEO ID'].astype(int)
Census_df['ZCTA5CE10'] = Census_df['ZCTA5CE10'].astype(int)
LRFs_df['GRID_ID'] = LRFs_df['GRID_ID'].astype(int)

Census_df = pd.concat([Census_df_org,missing_df])
Merged1 = Census_df.merge(LRFs_df, left_on='GEO ID', right_on='GRID_ID')
#grouped = Merged1['BIPD/PIP'].groupby(Merged1['ZCTA5CE10']) - this was for testing b4

def wavg(group, avg_name, weight_name):
    d = group[avg_name]
    w = group[weight_name]
    try:
        return round((d * w).sum() / w.sum(),3)
    except ZeroDivisionError:
        return np.NaN

lrf_0_Min = Merged1.groupby('ZCTA5CE10')['lrf_0'].min()
lrf_0_Max = Merged1.groupby('ZCTA5CE10')['lrf_0'].max()
lrf_1_Min = Merged1.groupby('ZCTA5CE10')['lrf_1'].min()
lrf_1_Max = Merged1.groupby('ZCTA5CE10')['lrf_1'].max()
lrf_2_Min = Merged1.groupby('ZCTA5CE10')['lrf_2'].min()
lrf_2_Max = Merged1.groupby('ZCTA5CE10')['lrf_2'].max()
lrf_3_Min = Merged1.groupby('ZCTA5CE10')['lrf_3'].min()
lrf_3_Max = Merged1.groupby('ZCTA5CE10')['lrf_3'].max()


lrf_0_wgt_avg = Merged1.groupby('ZCTA5CE10').apply(wavg, 'lrf_0', 'POP10')
lrf_1_wgt_avg = Merged1.groupby('ZCTA5CE10').apply(wavg, 'lrf_1', 'POP10')
lrf_2_wgt_avg = Merged1.groupby('ZCTA5CE10').apply(wavg, 'lrf_2', 'POP10')
lrf_3_wgt_avg = Merged1.groupby('ZCTA5CE10').apply(wavg, 'lrf_3', 'POP10')

summary = pd.concat([lrf_0_Min, lrf_0_wgt_avg, lrf_0_Max, \
                                 lrf_1_Min, lrf_1_wgt_avg, lrf_1_Max, \
                                 lrf_2_Min, lrf_2_wgt_avg, lrf_2_Max, \
                                 lrf_3_Min, lrf_3_wgt_avg, lrf_3_Max], axis=1)
#summary = pd.DataFrame(data=dict(s1=BI_Min, s2=BI_wgt_avg, s3=BI_Max, \
#                                 s4=Comp_Min, s5=Comp_wgt_avg, s6=Comp_Max, \
#                                 s7=Coll_Min, s8=Coll_wgt_avg, s9=Coll_Max, \
#                                 s10=Med_Min, s11=Med_wgt_avg, s12=Med_Max))
summary.columns = ['lrf_0 Min','lrf_0 Avg', 'lrf_0 Max', 'lrf_1 Min', 'lrf_1 Avg', 'lrf_1 Max', \
'lrf_2 Min','lrf_2 Avg', 'lrf_2 Max', 'lrf_3 Min', 'lrf_3 Avg', 'lrf_3 Max']

#summary1 = pd.concat([BI_Min, BI_wgt_avg, BI_Max, Comp_Min, Comp_wgt_avg, Comp_Max, Coll_Min, Coll_wgt_avg, Coll_Max], axis=1)
#summary1.columns = ['BI Min','BI Avg', 'BI Max', 'COMP Min','COMP Avg', 'COMP Max', 'COLL Min', 'COLL Avg', 'COLL Max']

"""Calculate the Min, Max, Average for each CC Zip"""

CC_Zip = pd.merge(Zone_df, summary, left_on="Zip_Code", right_index=True, how='left')

"""Calculate the Min, Max, Average (weighted by Pop and Exp) for each CC Zone"""

BI_Min_Zone = CC_Zip.groupby('Zone')['BI/PD Min'].min()
BI_Max_Zone = CC_Zip.groupby('Zone')['BI/PD Max'].max()
Comp_Min_Zone = CC_Zip.groupby('Zone')['COMP Min'].min()
Comp_Max_Zone = CC_Zip.groupby('Zone')['COMP Max'].max()
Coll_Min_Zone = CC_Zip.groupby('Zone')['COLL Min'].min()
Coll_Max_Zone = CC_Zip.groupby('Zone')['COLL Max'].max()
Med_Min_Zone = CC_Zip.groupby('Zone')['Med/Pip Min'].min()
Med_Max_Zone = CC_Zip.groupby('Zone')['Med/Pip Max'].max()


BI_wgt_Pop = CC_Zip.groupby('Zone').apply(wavg, 'BI/PD Avg', 'Population')
Comp_wgt_Pop = CC_Zip.groupby('Zone').apply(wavg, 'COMP Avg', 'Population')
Coll_wgt_Pop = CC_Zip.groupby('Zone').apply(wavg, 'COLL Avg', 'Population')
Med_wgt_Pop = CC_Zip.groupby('Zone').apply(wavg, 'Med/Pip Avg', 'Population')

BI_wgt_Exp = CC_Zip.groupby('Zone').apply(wavg, 'BI/PD Avg', 'Wrt_Exp')
Comp_wgt_Exp = CC_Zip.groupby('Zone').apply(wavg, 'COMP Avg', 'Wrt_Exp')
Coll_wgt_Exp = CC_Zip.groupby('Zone').apply(wavg, 'COLL Avg', 'Wrt_Exp')
Med_wgt_Exp = CC_Zip.groupby('Zone').apply(wavg, 'Med/Pip Avg', 'Wrt_Exp')

#CC_Zone = pd.DataFrame(data=dict(s1=BI_Min_Zone, s2=BI_wgt_Pop, s3=BI_wgt_Exp, s4=BI_Max_Zone, \
#                                 s5=Comp_Min_Zone, s6=Comp_wgt_Pop, s7=Comp_wgt_Exp, s8=Comp_Max_Zone, \
#                                 s9=Coll_Min_Zone, s10=Coll_wgt_Pop, s11=Coll_wgt_Exp, s12=Coll_Max_Zone))
#CC_Zone.columns = ['BI Min','BI Avg Pop', 'BI Avg Exp', 'BI Max', 'COMP Min','COMP Avg Pop', 'COMP Avg Exp', 'COMP Max', 'COLL Min','COLL Avg Pop', 'COLL Avg Exp', 'COLL Max']
#2 lines of code above give the wrong order

#CC_Zone = pd.DataFrame(data=dict(s1=BI_Min_Zone, s2=BI_wgt_Pop, s3=BI_Max_Zone, \
#                                 s4=Comp_Min_Zone, s5=Comp_wgt_Pop, s6=Comp_Max_Zone, \
#                                 s7=Coll_Min_Zone, s8=Coll_wgt_Pop, s9=Coll_Max_Zone))
#CC_Zone.columns = ['BI Min','BI Avg Pop', 'BI Max', 'COMP Min','COMP Avg Pop', 'COMP Max', 'COLL Min','COLL Avg Pop', 'COLL Max']
#2 lines of code however give the right order.Similarly there's no problem with the weighted avg by zip above. WEIRD

Zone_Pop = pd.concat([BI_Min_Zone, BI_wgt_Pop, BI_wgt_Exp, BI_Max_Zone, \
                                 Coll_Min_Zone, Coll_wgt_Pop, Coll_wgt_Exp, Coll_Max_Zone, \
                                 Comp_Min_Zone, Comp_wgt_Pop, Comp_wgt_Exp, Comp_Max_Zone, \
                                 Med_Min_Zone, Med_wgt_Pop, Med_wgt_Exp, Med_Max_Zone], axis=1)
Zone_Pop.columns = ['BI/PD Min', 'BI/PD Avg Pop', 'BI/PD Avg Exp', 'BI/PD Max', \
'COLL Min', 'COLL Avg Pop', 'COLL Avg Exp', 'COLL Max', \
'COMP Min', 'COMP Avg Pop', 'COMP Avg Exp', 'COMP Max', \
'Med/Pip Min', 'Med/Pip Avg Pop', 'Med/Pip Avg Exp', 'Med/Pip Max']

#Zone_Exp = pd.concat([BI_Min_Zone, BI_wgt_Exp, BI_Max_Zone, \
#                                 Comp_Min_Zone, Comp_wgt_Exp, Comp_Max_Zone, \
#                                 Coll_Min_Zone, Coll_wgt_Exp, Coll_Max_Zone, \
#                                 Med_Min_Zone, Med_wgt_Exp, Med_Max_Zone], axis=1)
#Zone_Exp.columns = ['BI/PD Min', 'BI/PD Avg Exp', 'BI/PD Max', \
#'COMP Min', 'COMP Avg Exp', 'COMP Max', \
#'COLL Min', 'COLL Avg Exp', 'COLL Max', \
#'Med/Pip Min', 'Med/Pip Avg Exp', 'Med/Pip Max']


"""Write results to Excel file"""

#using xlsxwriter

writer = pd.ExcelWriter(r'T:\ACT_PID\Pcm_brm\Misc\0.01 SF_LRFs\TT_Test\SF_%s_%s_LRFs_%s.xlsx' %(state, line, dt1), engine='xlsxwriter')
#LRFs_df.to_excel(writer,'SF-LRFs', index=False)
#Census_df_org.to_excel(writer,'Census', index=False)
CC_Zip.to_excel(writer,'CC-Zip', index=False)
Zone_Pop.to_excel(writer,'Zone-Pop-Exp')
#Zone_Exp.to_excel(writer,'Zone-Exp')
#CC_Zone1.to_excel(writer,'CC-Zone1')

pd.formats.format.header_style = None

wb  = writer.book
wsZip = writer.sheets['CC-Zip']
wsPop = writer.sheets['Zone-Pop-Exp']
#wsExp = writer.sheets['Zone-Exp']
#wsLRFs = writer.sheets['SF-LRFs']
#wsCensus = writer.sheets['Census']

#wsZone1 = writer.sheets['CC-Zone1']

#wsLRFs.set_column("A:A", 11)

No_border = wb.add_format({'border':0})
Lborder = wb.add_format({'left':1})
wsZip.set_column('A:W', None, No_border)
wsPop.set_column('A:W', None, No_border)
wsZip.freeze_panes('H2')
wsPop.freeze_panes('B2')
#wsExp.set_column('A:W', None, No_border)

format1 = wb.add_format({'bold': True, 'font_color': 'red'})
#wsZip.conditional_format('G2:G900', {'type':'formula', 'criteria':'=abs(H2/G2-1)>0.1', 'format':format1})                                        
#wsPop.conditional_format('B2:B900', {'type':'formula', 'criteria':'=abs(min(C2,D2)/B2-1)>0.1', 'format':format1})
from xlsxwriter.utility import xl_range

Rows = len(CC_Zip) + 1
Cols= len(CC_Zip.columns)

#range_min = xl_range(1, Cols-12, Rows, Cols-12)
#range_avg = xl_range(1, Cols-11, Rows, Cols-11)
#range_max = xl_range(1, Cols-10, Rows, Cols-10)

#wsZip.conditional_format(range_min, {'type':'cell', 'criteria':'==', 'value':'=abs(range_avg/range_min-1)>0.1', 'format':format1})
#for n in range(7, Cols, 3):
#    range0 = xl_range(0, n, 0, n)
#    #range00 = xl_range(Rows + 1, n, 1048576, n)
#    wsZip.set_column(range0, None, Lborder)
#    #wsZip.set_column(range00, None, No_border)
#
#    
#
#range1 = xl_range(1, Cols-12, Rows, Cols-12)
#range2 = xl_range(1, Cols-10, Rows, Cols-10)
#range3 = xl_range(1, Cols-9, Rows, Cols-9)
#range4 = xl_range(1, Cols-7, Rows, Cols-7)
#range5 = xl_range(1, Cols-6, Rows, Cols-6)
#range6 = xl_range(1, Cols-4, Rows, Cols-4)
#range7 = xl_range(1, Cols-3, Rows, Cols-3)
#range8 = xl_range(1, Cols-1, Rows, Cols-1)
#k = 2
#
##wb.define_name('Icol', '=CC-Zip!I2')
#wsZip.conditional_format(range1, {'type':'formula', 'criteria':'=abs(H2/I2-1)>0.1', 'format':format1})
#wsZip.conditional_format(range3, {'type':'formula', 'criteria':'=abs(K2/L2-1)>0.1', 'format':format1})
#wsZip.conditional_format(range5, {'type':'formula', 'criteria':'=abs(N2/O2-1)>0.1', 'format':format1})
#wsZip.conditional_format(range7, {'type':'formula', 'criteria':'=abs(Q2/R2-1)>0.1', 'format':format1})
#wsZip.conditional_format(range2, {'type':'formula', 'criteria':'=abs(J2/I2-1)>0.1', 'format':format1})
#wsZip.conditional_format(range4, {'type':'formula', 'criteria':'=abs(M2/L2-1)>0.1', 'format':format1})
#wsZip.conditional_format(range6, {'type':'formula', 'criteria':'=abs(P2/O2-1)>0.1', 'format':format1})
#wsZip.conditional_format(range8, {'type':'formula', 'criteria':'=abs(S2/R2-1)>0.1', 'format':format1})

#wsZip.conditional_format(range8, {'type':'formula', 'criteria':'=abs('S'+str(k)+'/R2-1)>0.1', 'format':format1})
#xl_range(0, 0, len(CC_Zip), len(CC_Zip.columns))
                                        
writer.save()

#using win32com.client
#excel.Visible = False

excel = win32.gencache.EnsureDispatch('Excel.Application')

#excel.ScreenUpdating = False

wb = excel.Workbooks.Open(r'T:\ACT_PID\Pcm_brm\Misc\0.01 SF_LRFs\TT_Test\SF_%s_%s_LRFs_%s.xlsx' %(state, line, dt1))
wb.Visible = False
wb.ScreenUpdating = False

wsZip = wb.Worksheets('CC-Zip')
wsZip.ListObjects.Add().TableStyle = "TableStyleMedium2"
wsPop = wb.Worksheets('Zone-Pop-Exp')
wsPop.ListObjects.Add().TableStyle = "TableStyleMedium2"
wsZip.Columns("A:F").AutoFit()
wsZip.Columns("G:S").ColumnWidth = 9
wsPop.Columns("A:Q").ColumnWidth = 10
wsZip.Range("E:E").HorizontalAlignment = win32.constants.xlCenter
wsPop.Columns(1).HorizontalAlignment = win32.constants.xlCenter
wsZip.Rows('1:1').WrapText = True
wsPop.Rows('1:1').WrapText = True 
wsZip.Rows.VerticalAlignment = win32.constants.xlTop
wsPop.Rows.VerticalAlignment = win32.constants.xlTop

wsPop.Columns(5).Borders(10).LineStyle=1
Rows = len(CC_Zip) + 1
Cols= len(CC_Zip.columns)

for m in range(8, Cols, 3):
    wsZip.Range(wsZip.Cells(1,m-1), wsZip.Cells(Rows,m-1)).Borders(10).LineStyle=1
    for l in range(2, Rows - 1):
        if abs(wsZip.Cells(l,m).Value/wsZip.Cells(l,m+1).Value-1)>0.1:
            wsZip.Cells(l,m).Font.Color = -16776961
            wsZip.Cells(l,m).Font.Bold = True
        if abs(wsZip.Cells(l,m+2).Value/wsZip.Cells(l,m+1).Value-1)>0.1:
            wsZip.Cells(l,m+2).Font.Color = -16776961
            wsZip.Cells(l,m+2).Font.Bold = True

Pop_Rows = len(Zone_Pop) + 1 
Pop_Cols= len(Zone_Pop.columns)
for n in range(2, Pop_Cols, 4):
    wsPop.Range(wsPop.Cells(1,n-1), wsPop.Cells(Pop_Rows,n-1)).Borders(10).LineStyle=1
    for k in range (2, Pop_Rows + 1):
        if abs(wsPop.Cells(k,n).Value/min(wsPop.Cells(k,n+1).Value,wsPop.Cells(k,n+2).Value)-1)>0.1:
            wsPop.Cells(k,n).Font.Color = -16776961
            wsPop.Cells(k,n).Font.Bold = True
        if abs(wsPop.Cells(k,n+3).Value/max(wsPop.Cells(k,n+1).Value,wsPop.Cells(k,n+2).Value)-1)>0.1:
            wsPop.Cells(k,n+3).Font.Color = -16776961
            wsPop.Cells(k,n+3).Font.Bold = True        

#wsPop.Cells(2,2).Font.Color = -16776961

#wsZip.Columns("G2:G900").Select
#formula1 = "=abs(min(C2,D2)/B2-1)>0.1"
#formula1 = "=0.711"
#wsZip.Columns("G:G").FormatConditions.Add(win32.constants.xlExpression, formula1)
#wsZip.Columns("G2:G900").FormatConditions(1).Font.Bold = True
#wsZip.Columns("G2:G900").FormatConditions(1).Font.Color = 5296274
#wsZip.Columns("G2:G900").FormatConditions(1).StopIfTrue = False

#writer.save() #use this to save the writer file created by xlsxwriter that includes the formatting done by win32

wb.Save()
#excel.ScreenUpdating = True
#wb.SaveAs(r'T:\ACT_PID\Pcm_brm\Misc\0.01 SF_LRFs\TT_Test\SF_%s_%s_LRFs_%s.xlsx' %(state, line, dt1), FileFormat=50)
#wb.Close()
#import os
#os.remove(r'T:\ACT_PID\Pcm_brm\Misc\0.01 SF_LRFs\TT_Test\SF_%s_%s_LRFs_%s.xlsx' %(state, line, dt1))

#excel.Application.Quit()


#------------TO DO / ENHANCEMENT-------------
#Change to read Excel file (Master_SF_LRFs.xlsx) to get missing_df
#Change to read LRFs_df straight from website
    #https://b2b.statefarm.com/b2b/roles/doi_public/DE_Auto_LRF_Effective_2017-04-17.pdf
#Change to not create the MED columns - some states only have LFRs for BI, Comp & Coll
#Check SF LRFs file, the first column name is different from st to st
#Handle error when calculating weighted average - this will occur when we have no LRFs for zip code

"""
Reference:

Learning Python, 5th Ed
Programming Python, 4th Ed
Python for Data Analysis, 1st Ed
http://pythonexcels.com/
http://pbpython.com/archives.html
https://pandas.pydata.org/pandas-docs/stable/
https://www.dataquest.io/blog/python-pandas-databases/
https://gist.github.com/hunterowens/08ebbb678255f33bba94
https://security.openstack.org/guidelines/dg_parameterize-database-queries.html
https://blogs.msdn.microsoft.com/cdndevs/2015/03/11/python-and-data-sql-server-as-a-data-source-for-python-applications/  

"""
#-------------NOTES------------------
#Modules that work well with Excel: ExcelWriter (including in pd), win32com.client, XlsxWriter, & openpyxl
#The last 2 appear to be popular
#2nd: http://pythonexcels.com/ & http://nbviewer.jupyter.org/github/pybokeh/jupyter_notebooks/blob/master/xlwings/Excel_Formatting.ipynb
#3rd: https://xlsxwriter.readthedocs.io/index.html
#4th: https://openpyxl.readthedocs.io/en/default/ -was said mainly for Excel 2007

#for pdf: https://www.binpress.com/tutorial/manipulating-pdfs-with-python/167


#-----------------------Test--------------

##LONG way to create the df for each group then merge one at a time        
#BI = pd.DataFrame(Merged1.groupby("ZCTA5CE10").apply(wavg, "BIPD/PIP", "POP10"))
#COMP = pd.DataFrame(Merged1.groupby("ZCTA5CE10").apply(wavg, "COMP", "POP10"))
#COLL = pd.DataFrame(Merged1.groupby("ZCTA5CE10").apply(wavg, "COLL", "POP10"))
#
#Merged2 = pd.merge(Zone_df, BI, left_on="Zip_Code", right_index=True)
#Merged2 = pd.merge(Merged2, COMP, left_on="Zip_Code", right_index=True)
#Merged2 = pd.merge(Merged2, COLL, left_on="Zip_Code", right_index=True)
#
##way to create ONE df to calculate Min, Max for each coverage, but it only works for standard agregate functions
#f = {'BIPD/PIP': ['min','max'],'COMP': ['min','max'], 'COLL': ['min', 'max']}
#f = {'BIPD/PIP': [['min'],['max']],'COMP': ['min','max'], 'COLL': ['min', 'max']}
#MinMax = Merged1.groupby('ZCTA5CE10').agg(f)
#
#summary = pd.DataFrame(data=dict(s1=MinMax, s2=BI_wgt_avg)) #<~~ this will fails, so the onlyway to do is merge them up
#
##Misc
#Census_df = pd.read_csv('Census.csv')
#Zone_df = pd.read_csv('Zone.csv')
#LRFs_df['GRID_ID'][0:4]
#LRFs_df.columns
#Zone_df[0:4]
#Census_df['Lat'] = str(Census_df['INTPTLAT10'][0:3])
#
#Merged1['Sum'] = grouped.sum()
#
#grouped.sum()
#
#Comp = COMP.to_frame().reset_index()
#
#df = DataFrame({'key1' : ['a', 'a', 'b', 'b', 'a'], 
#    'key2' : ['one', 'two', 'one', 'two', 'one'],
#    'data1' : np.random.randn(5),
#    'data2' : np.random.randn(5)})
#
#
#people = DataFrame(np.random.randn(5, 5),
#     columns=['a', 'b', 'c', 'd', 'e'],
#    index=['Joe', 'Steve', 'Wes', 'Jim', 'Travis'])
#
#people.ix[2:3,['b','c']] = np.nan
#
#
#st_code = states.get(state)[1]
#
#st_code = str(states.get(state)[1]) 
##without str is ok too, since st_code is already string
##as for query, Where a.STATEFP10 = %s is the same as Where a.STATEFP10 = '%s'
##probably because on SQL Server, using Where a.STATEFP10 = '04' or Where a.STATEFP10 = 4 gives the same result
#
#writer = pd.ExcelWriter('SF_%s_%s_LRFs_%s.xlsx' %(state, line, dt1))
##LRFs_df.to_excel(writer,'SF_LRFs', index=False)
#CC_Zip.to_excel(writer,'Zip', index=False)
#CC_Zone.to_excel(writer,'Zone')
#writer.save()
#
##Border = wb.add_format({'border':1})
##L_border = wb.add_format({'left':2})
##R_border = wb.add_format({'right':1})
#
##wsZip.set_column('G:R', 9, Border)
##wsZip.set_column('G:G', 9, L_border)
##wsZip.set_column('J:J', 9, L_border)
##wsZip.set_column('M:M',9, L_border)
##wsZip.set_column('P:P', 9, L_border)
#
#
##using Pandas
#writer = pd.ExcelWriter('SF_%s_%s_LRFs_%s.xlsx' %(state, line, dt1))
##LRFs_df.to_excel(writer,'SF_LRFs', index=False)
#CC_Zip.to_excel(writer,'Zip', index=False)
#CC_Zone.to_excel(writer,'Zone')
#writer.save()
#
##the below needs to try further to see if there's a shorter way to loop through sheets and format
#for sh in ('wsZip','wsZone'):
#    sh.Rows("1:1").WrapText = True
