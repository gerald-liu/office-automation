# %%
import pandas as pd

# %%
def parse_dict_str(s):
    if s == 'nan':
        return {}

    # Remove the curly braces from the string
    s = s.strip('{}')

    # Split the string into key-value pairs
    pairs = s.split(',')

    d = {}
    for pair in pairs:
        if ':' in pair:
            key, value = pair.split(':')
            d[key] = value
    return d


# %%
in_filename = 'toLuke销售分成修改版.xlsx'
out_filename_prefix = '销售分成修改版'
out_ext = '.xlsx'
col_filter = '新sales数据' # every value is '{key:value}'
vals_to_drop = ['PWM']
col_to_drop = ['原sales数据', '对比']

# %%
df_in = pd.read_excel(in_filename)

# %%
df = df_in.drop(columns=col_to_drop)

# %%
df['dicts'] = df[col_filter].astype(str).apply(parse_dict_str)

# %%
list_dup_keys = [set(d.keys()) for d in df['dicts']]
list_keys = list(set().union(*list_dup_keys)) # 销售姓名，无重复

for val in vals_to_drop:
    list_keys.remove(val)

# %%
dict_df_keys = {}

list_key_df_numrows = [] # 销售客户数量

for key in list_keys:
    mask = df['dicts'].apply(lambda x: key in x.keys())
    df_keys = df[mask].drop(columns='dicts')
    dict_df_keys[key] = df_keys

    list_key_df_numrows.append(df_keys.shape[0])

# %%
df_key_list = pd.DataFrame({
    '姓名': list_keys,
    '客户数量': list_key_df_numrows
})

df_key_list.sort_values(by='客户数量', ascending=False, inplace=True)
df_key_list.to_excel('销售列表.xlsx', header=True, index=False)

# %%
for key, df_out in dict_df_keys.items():
    out_filename = f'{out_filename_prefix}_{key}{out_ext}'
    df_out.to_excel(out_filename, index=False)
