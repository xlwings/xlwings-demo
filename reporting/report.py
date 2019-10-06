import os
import pandas as pd
from PIL import Image
from matplotlib.figure import Figure
from xlwings_reports import create_report  # not part of the open-source xlwings package

fig = Figure(figsize=(4, 3))
ax = fig.add_subplot(111)
ax.plot([1, 2, 3, 4, 5])

perf_data = pd.DataFrame(index=['r1', 'r1'],
                         columns=['c0', 'c1'],
                         data=[[1., 2.], [3., 4.]])

wb = create_report('template1.xlsx',
                   'output.xlsx',
                   perf=0.12 * 100,
                   perf_data=perf_data,
                   logo=Image.open(os.path.abspath('xlwings.jpg')),
                   fig=fig)
