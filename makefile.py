import pandas as pd
from pandas import DataFrame

class MakeFile:
    
    dataFile="5月data1.xlsx"
    baseFile="Webビデオ会議統計テンプレート.xlsx"
    ex_File="new3.xlsx"
    df1=DataFrame()
    df2=DataFrame()
    df=DataFrame()
    dict={}
    
    def readFile(self, dataFile, baseFile):
        self.df1=pd.read_excel(dataFile,sheet_name=0)
        self.df2=pd.read_excel(baseFile,sheet_name=1)

    def set_dict(self):
        self.dict={k:[v1,v2] for k,v1,v2 in zip(self.df2["メール"], self.df2["部署"], self.df2["名前"])}

    def searchInfo(self):
        column_email = [k for k in self.df1["email"]]
        column_department = []
        column_name = []
        for k in self.df1["email"]:
            if self.dict.get(k):
                column_department.append(self.dict.get(k)[0])
                column_name.append(self.dict.get(k)[1])
            else:
                column_department.append("    ")
                column_name.append("    ")
        return column_email,column_department,column_name
    
    
    def set_dataFrame(self):
        column_email,column_department,column_name=self.searchInfo()
        column_number = self.df1["总会议数"]
        df = pd.DataFrame(list(zip(column_email, column_department, column_name, column_number)),columns=["email","department","name","number"])
        self.df = df.sort_values("number", ascending=(False))
        return self.df

    def outputFile(self,dataframe,ex_File):
        dataframe.to_excel(ex_File, sheet_name='Sheet1')
        
        
        
makeFile=MakeFile()
makeFile.readFile(makeFile.dataFile,makeFile.baseFile)
makeFile.set_dict()
makeFile.set_dataFrame()
makeFile.outputFile(makeFile.df, makeFile.ex_File)