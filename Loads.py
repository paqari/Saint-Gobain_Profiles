import os
import pandas as pd

pd.set_option('display.max_columns', 20)

lc = pd.read_excel('Cb.xlsx', 'Base')

print(lc.tail(3))