import xlwings as xw
import time
import pandas as pd
import numpy as np
from functools import lru_cache

@xw.sub  # only required if you want to import it or run it via UDF Server
def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"


@xw.func
@xw.ret('raw')
def return_raw():
    df = pd.DataFrame(data=np.random.randn(1000, 1000),
                      index=pd.date_range('2019-01-01', freq='D', periods=1000))
    arr = df.to_numpy()
    return arr


@xw.func
@xw.arg('x', 'raw')
def read_raw(x):
    return x


@lru_cache()
@xw.func
def slow():
    time.sleep(5)
    return 'done'