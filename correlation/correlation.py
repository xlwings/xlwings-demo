import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
import seaborn as sns


@xw.func
@xw.arg('x', pd.DataFrame)
# @xw.ret(expand='table')  # use this if your version of Excel doesn't have dynamic arrays
def CORREL2(x):
    return x.corr()


@xw.func
@xw.arg('corr', pd.DataFrame)
def corr_plot(corr, caller):
    ax = sns.heatmap(corr, cmap='coolwarm', vmin=-1, vmax=1, linewidths=.5,
                     xticklabels=True, yticklabels=True)
    ax.tick_params(left=False, bottom=False)
    plt.yticks(rotation=0)
    plt.xticks(rotation=90)
    caller.sheet.pictures.add(ax.get_figure(),
                              top=caller.offset(row_offset=1).top,
                              left=caller.left,
                              name='CorrPlot',
                              update=True)
    return '<Corr Plot>'

