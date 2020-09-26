import pandas as pd

df1 = pd.DataFrame({
    'ix': [10, 22, 36, 49, 57],
    'b': ['a', 'b', 'c', None, 'd']
})
df2 = pd.DataFrame({'ix': [10, 22, 57], 'b': ['e', None, 'f']})
a = dict(zip(df1['ix'], df1['b']))
b = dict(zip(df2['ix'], df2['b']))
tmp_d = {}
tmp_d.update(a)
tmp_d.update(b)
df1 = pd.DataFrame(tmp_d.items()).rename(columns={0: 'ix', 1: 'b'})
print(df1)
