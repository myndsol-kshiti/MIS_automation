import pandas as pd

#load all files
#mis
old_mis = pd.read_excel(r"C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\Indus SSC Combined MIS -04-JULY-22 - Copy.xlsx",sheet_name="Data")
new_header = old_mis.iloc[0]
old_mis = old_mis[1:]
old_mis.columns = new_header
print("old mis loaded")
#phd combined
phd = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\Comp.-5Jul DONE.xlsb', engine='pyxlsb')
phd.rename(columns = {'Inward No':'Docket No.'}, inplace = True)
print("phd loaded")
#ssc batch
ssc_batch = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\SSC Batch.xlsx')
ssc_batch_rejection = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\Rejection Batch.xlsx')
ssc_batch_isupplier = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\Isupplier Batch.xlsx')
print("ssc batch loaded")
#eb-gst batch
eb_gst_batch = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\EB_GST.xlsx',sheet_name='SSC Batch')
eb_gst_rejection = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\EB_GST.xlsx',sheet_name='SSC Rejection')
eb_gst_hold = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\EB_GST.xlsx',sheet_name='OTC Hold Data')
print("eb gst batch loaded")

#grn_done
grn_done = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\GRN DONE 6 JUL 22.xlsx',sheet_name='GRN DATA')
print("grn done loaded")

#asn_file
asn = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\ASN-6 JUl.xlsb',engine='pyxlsb')
#eb-gst-mis
# eb_gst_mis = pd.read_excel()
# eb_gst_mis_docket = set()
# print("eb gst mis loaded")

#i-expense
# i_expense = pd.read_excel()
# i_expense_docket = set()
# print("i expense loaded")
#eastern mis
eastern_mis = pd.read_excel(r"C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\EASTERN MIS_5-JUL-22.xlsb",sheet_name='Data',engine='pyxlsb')
new_header1 = eastern_mis.iloc[0]
eastern_mis = eastern_mis[1:]
eastern_mis.columns = new_header1
print("eastern mis loaded")

#vlookup with phd and old mis

mis_phd_common = pd.merge(old_mis,
                          phd,
                          on ='Docket No.',
                          how ='inner')
common_docket_phd = mis_phd_common['Docket No.'].tolist()

#check for phd file

options1 = ['Hold']
options2 = ['OTC']
options3 = ['SCM']
options4 = ['User']

mis_phd_hold = mis_phd_common[mis_phd_common['Current Status'].isin(options1)]
mis_phd_otc = mis_phd_common[mis_phd_common['Current Status'].isin(options2)]
mis_phd_scm = mis_phd_common[mis_phd_common['Current Status'].isin(options3)]
mis_phd_user = mis_phd_common[mis_phd_common['Current Status'].isin(options4)]

#phd file inputs
set_hold = set(mis_phd_hold['Docket No.'])
set_otc = set(mis_phd_hold['Docket No.'])
set_scm = set(mis_phd_scm['Docket No.'])
set_user = set(mis_phd_user['Docket No.'])

#ssc batch combine
ssc_batch_docket = set(ssc_batch['InwardNo'])
ssc_batch_reject_docket = set(ssc_batch_rejection['InwardNo'])
ssc_batch_isupplier_docket = set(ssc_batch_isupplier['Unique Id'])

ssc_batch1 = ssc_batch_docket.union(ssc_batch_reject_docket)
ssc_batch = ssc_batch1.union(ssc_batch_isupplier_docket)

#eb-gst batch combine
eb_gst_batch_docket = set(eb_gst_batch['InwardNo'])
eb_gst_reject_docket = set(eb_gst_rejection['InwardNo'])
eb_gst_otc_docket = set(eb_gst_hold['InwardNo'])

eb_gst_batch1 = eb_gst_batch_docket.union(eb_gst_reject_docket)
eb_gst_batch = eb_gst_batch1.union(eb_gst_otc_docket)

#mis and ssc batch common docket
mis_docket = set(old_mis['Docket No.'])
ssc_batch_mis = mis_docket.intersection(ssc_batch)
print(len(ssc_batch_mis))

#mis and eb gst batch common docket
eb_gst_batch_mis = mis_docket.intersection(eb_gst_batch)
print(len(eb_gst_batch))

#mis and eb-gst mis common docket
# eb_gst_mis = mis_docket.intersection(eb_gst_mis_docket)
# print(len(eb_gst_batch))

#mis and i-expense common docket
# i_expense_mis = mis_docket.intersection(i_expense_docket)
# print(len(i_expense_mis))

#eastern and mis common docket
eastern_docket = set(eastern_mis['Docket No.'])
eastern_mis_docket = mis_docket.intersection(eastern_docket)
eastern_mis_docket = list(eastern_mis_docket)

eastern_mis_columns = eastern_mis[eastern_mis['Docket No.'].isin(eastern_mis_docket)]
print(len(eastern_mis_columns))
old_mis_for_eastern = old_mis[old_mis['Docket No.'].isin(eastern_mis_docket)]

#grn and mis combine
grn_done_docket = set(grn_done['Docket No.'])

#asn_filter
options_asn_status =['Pending with User','Pending with SCM']
options_asn_add = ['workflow approval pending with user','ASN Approval']

mis_asn1 = old_mis[old_mis['Status of Invoice'].isin(options_asn_status)]
mis_asn = mis_asn1[mis_asn1['Reason of Rejection/Hold'].notin(options_asn_add)]

#asn-mis
asn_docket = set(mis_asn['ASN Number'])
mis_invoice_num = set(old_mis['Invoice No.'])
mis_asn_docket = mis_invoice_num.intersection(asn_docket)
mis_asn_docket = list(mis_asn_docket)
#asn-status-docket
mis_asn2 = mis_asn[mis_asn['Docket No.'].isin(mis_asn_docket)]
list_asn_status_scm = ['0','APPROVED','#N/A']
list_asn_status_user1=['APPROVAL PENDING']
list_asn_status_user2 = ['CANCELLED']
asn_scm = mis_asn[mis_asn['Docket No.'].isin(list_asn_status_scm)]
asn_scm_doc = set(asn_scm['ASN Number'])
asn_user1 = mis_asn[mis_asn['Docket No.'].isin(list_asn_status_user1)]
asn_user1_doc = set(asn_user1['ASN Number'])
asn_user2 =  mis_asn[mis_asn['Docket No.'].isin(list_asn_status_user2)]
asn_user2_doc = set(asn_user2['ASN Number'])

master = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\Indus SSC Combined MIS -04-JULY-22 - Copy.xlsx',sheet_name="Data")
new_header = master.iloc[0]
master = master[1:]
master.columns = new_header
len_doc = len(master['Docket No.'])
print(len_doc)
for i in range(1,len_doc):
    #phd competency loop
    print(master['Docket No.'][i])
    if master['Docket No.'][i] in set_hold:
        master['Status of Invoice'][i] == "Hold at SSC"
        # master['Reason of Reason of Rejection/Hold'][i] ==
    elif master['Docket No.'][i] in set_otc:
        master['Status of Invoice'][i] == "Rejected to Partner"
        master['Reason of Rejection/Hold'][i] == "Rejected to Partner"
    elif master['Docket No.'][i] in set_scm:
        master['Status of Invoice'][i] == "Pending with SCM"
        master['Reason of Rejection/Hold'][i] == "Pending with GRN"
    elif master['Docket No.'][i] in set_user:
        master['Status of Invoice'][i] == "Pending with User"
        master['Reason of Rejection/Hold'][i] == "Pending with GRN"
    master1 = master
    for i in range(1, len_doc):
    #ssc_batch_loop
        print(master['Docket No.'][i])
        if master1['Docket No.'][i] in ssc_batch_mis:
            master1['Status of Invoice'][i] == "Pending with SSC"
            master1['Reason of Rejection/Hold'][i] == "Ready for SSC"
    master2 = master1
    for i in range(1, len_doc):
    #eb_gstBatch loop
        print(master['Docket No.'][i])
        if master2['Docket No.'][i] in eb_gst_batch_mis:
            master2['Status of Invoice'][i] == "Rejected to Partner"
            master2['Reason of Rejection/Hold'][i] == "Belong to EB-GST/already to EB GST"
master3 = master
    # eb-gst-mis-loop
    # if master['Docket No.'][i] == eb_gst_mis:
    #     master['Status of Invoice'][i] == "Rejected to Partner"
    #     master['Reason of Rejection/Hold'][i] == "Belong to EB-GST/already to EB GST"
    # i-expense loop
    #     if master['Docket No.'][i] == i_expense_mis:
    #         master['Status of Invoice'][i] == "Rejected to Partner"
    #         master['Reason of Rejection/Hold'][i] == "Belong to I expense"
for i in range(1, len_doc):
    # for j in range(1,len(eastern_mis_columns)):
    #eastern mis
        print(master3['Docket No.'])
        if master3['Docket No.'][i] in eastern_mis_docket:
            var_doc = list(master3['Docket No.'][i])
            mid_eastern =  eastern_mis_columns[eastern_mis_columns['Docket No.'].isin(var_doc)]
            master3['Status of Invoice'][i] == mid_eastern['Status of Invoice']
            master3['Reason of Rejection/Hold'][i] == mid_eastern['Reason of Rejection/Hold']
            master3['Additional Remarks'][i] == mid_eastern['Additional Remarks']

options_status = ['Due For Payment','Not Due for Payment','Validated']
option_add_remark = ['Circle_SSL']
eastern_mis_check = master3[master3['Status of Invoice'].isin(options_status)]
eastern_mis_check_2 = eastern_mis_check[eastern_mis_check['Additional Remarks'].isin(option_add_remark)]

mis_eastern_add_check = set(eastern_mis_check_2['Docket No.'])

master4 = master3
# eastern-mis-checks
for i in range(1,len(master4)):
    if master4['Docket No.'][i] in mis_eastern_add_check:
        master4['Status of Invoice'][i] == 'Pending with SSC'
        master4['Reason of Rejection/Hold'][i] == 'WIP for Header'

for i in range(0,len(master4)):
    if master4['Docket No.'][i] in grn_done_docket:
        master4['Status of Invoice'][i] == 'Pending with SSC'
        master4['Reason of Rejection/Hold'][i] == 'WIP for Header'
for i in range(0,len(master4)):
     if master4['Docket No.'][i] in asn_scm_doc:
         master4['Status of Invoice'][i] == 'Pending with SCM'
         master4['Reason of Rejection/Hold'][i] == 'Pending with GRN'
     elif master4['Docket No.'][i] in asn_user1:
         master4['Status of Invoice'][i] == 'Pending with User'
         master4['Reason of Rejection/Hold'][i] == 'Pending for ASN Approval'
     elif master4['Docket No.'][i] in asn_user2:
         master4['Status of Invoice'][i] == 'Pending with User'
         master4['Reason of Rejection/Hold'][i] == 'ASN Cancelled/Rejected'







