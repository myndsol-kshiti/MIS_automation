import pandas as pd

df_allocation = pd.read_excel(r'C:\Users\hrithik.chauhan\Downloads\status_of_invoice_files\status_of_invoice_files\I-Supplier - Inflow.xlsx')
df_allocation.rename(columns = {'INWARD NO':'Docket No.'}, inplace = True)

print(df_allocation)



df_prev_mis = pd.read_excel(r'C:\Users\hrithik.chauhan\Downloads\status_of_invoice_files\status_of_invoice_files\test_mis_4july.xlsx')

print(df_prev_mis)
myset = set()


vlookup_common= pd.merge(df_allocation,
                    df_prev_mis,
                    on ='Docket No.',
                    how ='inner')
print(vlookup_common)

for i in range(0,len(vlookup_common)):
     myset.add(vlookup_common['Docket No.'][i])

l = []

# df_allocation_2 = pd.read_csv(url_2)
df_allocation_2 = pd.read_excel(r'C:\Users\hrithik.chauhan\Downloads\status_of_invoice_files\status_of_invoice_files\I-Supplier - Inflow.xlsx')
df_allocation_2.rename(columns = {'INWARD NO':'Docket No.'}, inplace = True)
for i in range(0,len(df_allocation_2['Docket No.'])):
  if df_allocation_2['Docket No.'][i] in myset:
    l.append(i)
for i in l:
 df_allocation_2.drop(i, inplace = True)

df_allocation_2.to_excel(r'C:\Users\hrithik.chauhan\Downloads\status_of_invoice_files\status_of_invoice_files\digital.xlsx')

print(df_allocation_2)
print(len(df_allocation_2))

