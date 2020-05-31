import pandas as pd
import datetime as datetime
from makefile import MakeFile as mf
from countTotalNum import CountTotalNum as ctn
from pandas import Series, DataFrame

class DateCountNum:

    dataFile="5月data2.xlsx"
    baseFile="Webビデオ会議統計テンプレート.xlsx"
    ex_File="new7.xlsx"
    df1=DataFrame()
    df2=DataFrame()
    df3=DataFrame()
    df=DataFrame()
    
    def set_dataFrame(self):
        column_email,column_department,column_name = mf.searchInfo(self)
        column_number = self.df1["总会议数"]
        column_userName = self.df1["userName"]
        column_date = self.df1["currentDay"]
        df = pd.DataFrame(list(zip(column_userName, column_name, column_department, column_email, column_date, column_number)), columns=["username","name","department","email","date","number"])
        self.df = df.sort_values("number", ascending=(False))
        return self.df
    

    def dayNumCnt(self,dataframe):
        i=0
        column=[]
        for department in ctn.list: 
            g = dataframe.groupby("department")
            d=g.get_group(department).sort_values("date")
            g1 = d.groupby("date").sum()
            df = pd.DataFrame(g1)
            self.df3 = pd.concat([self.df3,df],axis=1)
            i+=1
        self.df3.columns=ctn.list
        return self.df3    

    
dcn=DateCountNum()
mf.readFile(dcn,dcn.dataFile, dcn.baseFile)
mf.set_dict(dcn)
df = dcn.set_dataFrame()
dcn.df3 =dcn.dayNumCnt(df)
mf.outputFile(dcn,dcn.df3,dcn.ex_File)
