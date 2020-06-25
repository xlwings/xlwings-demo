import numpy as np
import xlwings as xw


@xw.func
def mysum(x, y, z):
    return x + y + z


@xw.func
@xw.arg('x', np.array, ndim=2)
@xw.arg('y', np.array)
def myarraysum(x, y, z):
    return x + y + z