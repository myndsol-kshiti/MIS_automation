import pandas as pd

df_allocation = pd.read_excel(r'C:\Users\kshiti.sinha\Downloads\ALLOCATION(I-SUPPLIER)-06-JULY-2022.xlsx')
df_allocation.rename(columns = {'Unique Id':'Docket No.'}, inplace = True)

print(df_allocation)
# sheet_url_mis = "https://docs.google.com/spreadsheets/d/1ERYEFuGi5Nh_3eNcGbU8PCLq0F68PiTTMVN2mu3fPIg/edit#gid=2098102185"
# url_1 = sheet_url_mis.replace('/edit#gid=', '/export?format=csv&gid=')
# df_prev_mis = pd.read_csv(url_1)

# sheet_url_alloc = "https://docs.google.com/spreadsheets/d/1OjRo56LYnBWAdVg0qyX2x5swtJPOfrOJsBPGUjn-mis/edit#gid=0"
# url_2 = sheet_url_alloc.replace('/edit#gid=', '/export?format=csv&gid=')
# df_allocation = pd.read_csv(url_2)


df_prev_mis = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\I-Supplier - Inflow.xlsx')
df_allocation.rename(columns = {'INWARD NO':'Docket No.'}, inplace = True)
# new_header = df_prev_mis.iloc[0]
# df_prev_mis = df_prev_mis[1:]
# df_prev_mis.columns = new_header
print(df_prev_mis)
myset = set()
# df_allocation.rename(columns = {'INWARD NO':'Docket No.'}, inplace = True)

vlookup_common= pd.merge(df_allocation,
                    df_prev_mis,
                    on ='Docket No.',
                    how ='inner')
print(vlookup_common)

for i in range(0,len(vlookup_common)):
     myset.add(vlookup_common['Docket No.'][i])

l = []

# df_allocation_2 = pd.read_csv(url_2)
df_allocation_2 = pd.read_excel(r'C:\Users\kshiti.sinha\Downloads\ALLOCATION(I-SUPPLIER)-06-JULY-2022.xlsx')
df_allocation_2.rename(columns = {'Unique Id':'Docket No.'}, inplace = True)
for i in range(0,len(df_allocation_2['Docket No.'])):
  if df_allocation_2['Docket No.'][i] in myset:
    l.append(i)
for i in l:
 df_allocation_2.drop(i, inplace = True)

df_allocation_2.to_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\digital.xlsx')

print(df_allocation_2)
print(len(df_allocation_2))

