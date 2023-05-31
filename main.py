import numpy as np
import pandas as pd
import xlwings as xw
pd.options.display.max_rows = 1000
#pd.options.display.max_columns = 100


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.



class SeCtion:
    def __init__(self,a,b):
        """

        :param a Это или сторона для квад сечения или Диаметр для круглой:
        :param b это или стороная или толщина стенки трубы:
        """
        self.a=a#мм
        self.b=b#мм
        self.D=a
        self.t=b

    def _Mom_sopr_kv(self):
        W=(self.a*self.b**2)/6
        return W

    @property
    def Mom_sopr_trubi(self):
        """
        Момент сопротивления круглой трубы
        :return:
        """
        d=self.D-2*self.t
        a=d/self.D
        Wx=np.pi/32*(pow(self.D,4)-pow(d,4))/32*(1-pow(a,4))
        Wx=Wx/1000000000# Перевод в метры
        return Wx
    @property
    def S_trubi(self):
        S=np.pi*(self.D/2-self.t)**2
        S=S/1000000
        return S




# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    book = xw.books
    sheet = book.active.sheets
    sheet=sheet.active

    usilia:pd.DataFrame
    usilia = sheet.range("A1").options(pd.DataFrame, expand='table', index_col=True).value
    name=usilia.columns
    print(name.values)


    truba=SeCtion(720,12)
    A=truba.S_trubi
    Wx=Wy=truba.Mom_sopr_trubi




    for i in name:
        print(i.split()[0])
        usilia.rename(columns={i:i.split(",")[0]},inplace=True)
    print(usilia)
    usilia.reset_index(inplace=True)
    usilia.drop(columns=["Номер РСН",'Тип ЭЛЕМ',"ЭЛЕМ","СЕЧ"],inplace=True)


    usilia["G"]=pow((usilia["MZ"]/Wx)**2+(usilia["MY"]/Wy)**2,0.5)+usilia["N"]/A

    new_usilia=usilia.loc[usilia.agg(['idxmax', 'idxmin']).stack()].drop_duplicates()

    print(new_usilia)
    sheet.range("K1").options(index=False).value = new_usilia
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
