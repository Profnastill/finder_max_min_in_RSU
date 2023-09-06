# Программа для выборки Максмимальных РСУ из таблиц РСУ
# Включить excel  с таблицей РСУ


import numpy as np
import pandas as pd
import xlwings as xw
pd.options.display.max_rows = 1000
pd.options.display.max_rows = 1000
pd.options.display.max_columns = 100
pd.options.display.expand_frame_repr = False




if __name__ == '__main__':
    book = xw.books
    sheet = book.active.sheets
    sheet=sheet.active

    usilia:pd.DataFrame
    usilia = sheet.range("A1").options(pd.DataFrame, expand='table', index_col=True).value
    name=usilia.columns
    print(name.values)

    for i in name:
        print(i.split()[0])
        usilia.rename(columns={i:i.split(",")[0]},inplace=True)

    usilia.reset_index(inplace=True)
    usilia.drop(columns=['MK','НС',"КРТ","СТ","КС","Г"],inplace=True)
    #new_usilia=usilia.loc[usilia.agg(['N', 'MK']).stack()].drop_duplicates()
    print(usilia)
    usilia["MY"]=usilia["MY"].abs()
    usilia["MZ"] = usilia["MZ"].abs()
    #new_usilia=usilia.agg(N_max=("N",max),N_min=("N",min), My=("MY",max),Mz=("MZ",max))
    #print(new_usilia)
    #new_usilia=usilia.loc[usilia['N'].idxmax()].to_frame().reset_index().transpose()
    new_usilia=usilia.query("N == N.max()|N == N.min()|MY == MY.max()|MZ == MZ.max()")
    print(new_usilia)
    xlsheet = sheet

    xlsheet.range("R1").options(index=False).value = new_usilia



# See PyCharm help at https://www.jetbrains.com/help/pycharm/
