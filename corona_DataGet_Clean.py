# 下载源文件
import pandas as pd
import wget

url = 'https://covid19.who.int/WHO-COVID-19-global-table-data.csv'
path = 'corona.csv'
wget.download(url, path)

# 数据读取


csv_file = "corona.csv"
csv_data = pd.read_csv(csv_file, low_memory=False)  # 防止弹出警告
df = csv_df = pd.DataFrame(csv_data)
# 删除不需要的数据
df1 = df.drop(labels=['Cases - cumulative total per 100000 population',
                      'Cases - newly reported in last 7 days',
                      'Cases - newly reported in last 7 days per 100000 population',
                      'Deaths - cumulative total per 100000 population',
                      'Deaths - newly reported in last 7 days',
                      'Deaths - newly reported in last 7 days per 100000 population',
                      'Transmission Classification'], axis=1)
# 按国家筛选数据
Countries = ['Poland', 'Ukraine', 'Czechia', 'Romania', 'Sweden', 'Serbia', 'Austria', 'Hungary', 'Slovakia',
             'Bulgaria',
             'Croatia', 'Denmark', 'Greece', 'Lithuania', 'Slovenia', 'Republic of Moldova', 'Bosnia and Herzegovina',
             'Albania', 'North Macedonia', 'Latvia', 'Montenegro', 'Estonia', 'Norway', 'Kosovo[1]', 'Finland', 'Malta',
             'Iceland']
df2 = df1[df1['Name'].isin(Countries)]

df2.rename(columns={'Cases - cumulative total': 'Totalcase',
                    'Cases - newly reported in last 24 hours': 'newcase',
                    'Deaths - cumulative total': 'Totaldeath',
                    'Deaths - newly reported in last 24 hours': 'newdeath'}, inplace=True)
df2.loc['Row_sum'] = df2.apply(lambda x: x.sum())

Nordic = ['Sweden', 'Finland', 'Iceland', 'Denmark', 'Norway']
df3 = df2[df2['Name'].isin(Nordic)]
df3.loc['Nordic_sum'] = df3.apply(lambda x: x.sum())

print(df2)
print(df3)

with pd.ExcelWriter("new file.xlsx") as writer:
    df3.to_excel(writer, sheet_name="Nordic", index=False)  # first是第一张工作表名称
    df2.to_excel(writer, sheet_name="CEE", index=False)  # second是第二张工作表名称
