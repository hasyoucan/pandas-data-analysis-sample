import pandas as pd
from makefile import MakeFile as mf
from pandas import Series, DataFrame

class CountTotalNum:

    filename="new3.xlsx"
    ex_File1="new5.xlsx"
    ex_File2="new6.xlsx"
    df1=DataFrame()
    df2=DataFrame()
    df3=DataFrame()
    list=["事業部",
          "企画本部",
          "中文広告チーム",
          "システム部",
          "pop部",
          "管理本部（日本总裁办）",
          "日本財務部"]
    
    def readFile(self, filename):
        self.df1=pd.read_excel(filename)
    
    def departTotalNum(self):
        g = self.df1.groupby(["department"])
        g1=g.sum().sort_values("number",ascending= [False])
        self.df2 = pd.DataFrame(g1)
        return self.df2
      
    def top3_member(self):
        g = self.df1.groupby(["department"])
        i=0
        df = DataFrame()
        for department in self.list:
            g1=g.get_group(self.list[i]).sort_values("number",ascending= [False]).head(3)
            df = pd.DataFrame(g1)
            self.df3 = pd.concat([self.df3,df], axis=0)
            i+=1
        return self.df3
     
        
ct = CountTotalNum()
ct.readFile(ct.filename)
ct.df2 = ct.departTotalNum()
mf.outputFile(ct,ct.df2,ct.ex_File1)
ct.df3 =ct.top3_member()
mf.outputFile(ct,ct.df3,ct.ex_File2)