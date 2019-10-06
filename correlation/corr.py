import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
import seaborn as sns


@xw.func
@xw.arg('x', pd.DataFrame)
def CORREL2(x):
    return x.corr()


@xw.func
@xw.arg('corr', pd.DataFrame)
# @xw.ret(expand='table')  # use this if your version of Excel doesn't have dynamic arrays
def corr_plot(corr):
    wb = xw.Book.caller()
    ax = sns.heatmap(corr, vmin=-1, vmax=1, linewidths=.5, xticklabels=True, yticklabels=True)
    plt.yticks(rotation=0)
    plt.xticks(rotation=90)
    fig = ax.get_figure()
    wb.sheets.active.pictures.add(fig,
                                  top=wb.selection.top,
                                  left=wb.selection.left,
                                  name='CorrPlot',
                                  update=True)
    plt.close()
    return '<Corr Plot>'

