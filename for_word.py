from tushare import MailMerge
import pandas as pd
import numpy as np
import json

doc = MailMerge('input.docx')

d = {'xiaoshou_mianji': 30, 'xiaoshou_mianji_huanbi': 50}
doc.merge(**d)

data = np.array(np.random.randn(6)).reshape(3, 2)

columns = ['col_one', 'col_two']
df = pd.DataFrame(data, columns=columns)
j = json.loads(df.to_json(orient='records'))
doc.merge(col_one=j)

doc.write('output.docx')
