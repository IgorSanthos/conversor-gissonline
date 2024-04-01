import pandas as pd

try:
    dfGiss = pd.read_csv('giss.csv', encoding='latin1' )# encoding usa-se para  ler csv separado por virgula
    print(dfGiss.head(10))

except Exception as error:
    print("Ocorreu algum erro",error)