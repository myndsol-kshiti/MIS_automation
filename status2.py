import pandas as pd
from datetime import datetime,date
# import sys
# sys.stdout = open(r'C:\Users\kshiti.sinha\PycharmProjects\codes_MIS\console_data', 'w')
import datetime
import smtplib
from email.message import EmailMessage
from datetime import date,datetime

#load all files
#mis
now1 = datetime.now()
print("starting time:",datetime.now())
# time_df = pd.DataFrame(columns='Time Check')
# time_df['Time Stamp'] = datetime.now()
old_mis = pd.read_excel(r"C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\data (19).xlsx",sheet_name='Export')
print("old mis loaded",len(old_mis))
# new_header = old_mis.iloc[0]
# old_mis = old_mis[1:]
# old_mis.columns = new_header
# prev_day_mis = pd.read_excel(r"C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\test_mis_4july.xlsx")
mis_docket = set(old_mis['Docket No.'])
mis_invoice_num = set(old_mis['Invoice No.'])
# print("prev day mis loaded",len(prev_day_mis))
#phd combined
phd = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\Comp.-5Jul DONE.xlsb', engine='pyxlsb')
phd.rename(columns = {'Inward No':'Docket No.'}, inplace = True)
print("phd loaded",len(phd))
#ssc batch
ssc_batch = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\SSC Batch.xlsx')
ssc_batch_rejection = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\Rejection Batch.xlsx')
ssc_batch_isupplier = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\Isupplier Batch.xlsx')
print("ssc batch loaded",len(ssc_batch))
# #eb-gst batch
eb_gst_batch = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\EB_GST.xlsx',sheet_name='SSC Batch')
eb_gst_rejection = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\EB_GST.xlsx',sheet_name='SSC Rejection')
eb_gst_hold = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\EB_GST.xlsx',sheet_name='OTC Hold Data')
print("eb gst batch loaded",len(eb_gst_batch))
#
# #eb-gst-mis
# # eb_gst_mis = pd.read_excel()
# # eb_gst_mis_docket = set()
# # print("eb gst mis loaded")
#
# #i-expense
# # i_expense = pd.read_excel()
# # i_expense_docket = set()
# # print("i expense loaded")
# #eastern mis
eastern_mis = pd.read_excel(r"C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\EASTERN MIS_5-JUL-22.xlsb",sheet_name='Data',engine='pyxlsb')
new_header1 = eastern_mis.iloc[0]
eastern_mis = eastern_mis[1:]
eastern_mis.columns = new_header1
print("eastern mis loaded",len(eastern_mis))
#
# #grn_done
grn_done = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\GRN DONE 6 JUL 22.xlsx',sheet_name='GRN DATA')
print("grn done loaded",len(grn_done))
#
# #asn_file
asn = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\ASN-6 JUl.xlsb',engine='pyxlsb')
print("asn loaded")
# #open invoice report
open_invoice = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\Open Invoice Report(05-07-2022)_AUTOMATION REPORT.xlsx')
new_header2 = open_invoice.iloc[0]
open_invoice = open_invoice[1:]
open_invoice.columns = new_header2
print("open invoice loaded",len(open_invoice))

# #ssc mis combined
ssc_mis_combined = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\SSC MIS Tracker 05-Jul-2022.xlsb',sheet_name='DATA',engine='pyxlsb')
print("ssc mis combined loaded",len(ssc_mis_combined))

#car report
car = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\CAR Report 06 Jul.xlsx')
print("car report loaded",len(car))

paid = pd.read_excel(r'C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\Paid Report.xlsx')
print("paid file loaded",len(paid))

print("LOADED ALL FILES")
print("loading files time:",datetime.now())
# #vlookup with phd and old mis
#
mis_phd_common = pd.merge(old_mis,
                          phd,
                          on ='Docket No.',
                          how ='inner')
common_docket_phd = mis_phd_common['Docket No.'].tolist()
print("common docket phd and mis",len(common_docket_phd))
#
# #check for phd file
#
options1 = ['Hold']
#2
options2 = ['OTC']
#58
#getting 39
options3 = ['SCM']
#74
options4 = ['User']
#no user

mis_phd_hold = mis_phd_common[mis_phd_common['Current Status'].isin(options1)]
print("hold",len(mis_phd_hold))
mis_phd_otc = mis_phd_common[mis_phd_common['Current Status'].isin(options2)]
print("otc",len(mis_phd_otc))
mis_phd_otc.to_excel("mis_phd_otc.xlsx")
mis_phd_scm = mis_phd_common[mis_phd_common['Current Status'].isin(options3)]
print("scm",len(mis_phd_scm))
mis_phd_user = mis_phd_common[mis_phd_common['Current Status'].isin(options4)]
print("user",mis_phd_user)

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
ssc_batch_mis = mis_docket.intersection(ssc_batch)
print("ssc and mis common",len(ssc_batch_mis))

#mis and eb gst batch common docket
eb_gst_batch_mis = mis_docket.intersection(eb_gst_batch)
print("eb gst and mis common",len(eb_gst_batch))

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
print("eastern and mis common",len(old_mis_for_eastern))
#grn and mis combine
grn_done_docket = set(grn_done['Docket No.'])

#asn_filter
options_asn_status =['Pending with User','Pending with SCM']
options_asn_add = ['workflow approval pending with user','ASN Approval']

mis_asn1 = old_mis[old_mis['Status of Invoice'].isin(options_asn_status)]
mis_asn = mis_asn1[mis_asn1['Reason of Rejection/Hold'].isin(options_asn_add) == False]
mis_asn_invoice = set(mis_asn['Invoice No.'])
mis_asn_invoice = list(mis_asn_invoice)

# asn-mis
asn_docket = set(mis_asn['Docket No.'])
mis_asn_docket = mis_invoice_num.intersection(asn_docket)
mis_asn_docket = list(mis_asn_docket)
#asn-status-docket
mis_asn2 = asn[asn['ASN Number'].isin(mis_asn_invoice)]
list_asn_status_scm = ['0','APPROVED','#N/A']
list_asn_status_user1=['APPROVAL PENDING']
list_asn_status_user2 = ['CANCELLED']
asn_scm = mis_asn2[mis_asn2['ASN Status'].isin(list_asn_status_scm)]
asn_scm_doc = set(asn_scm['ASN Number'])
asn_user1 = mis_asn2[mis_asn2['ASN Status'].isin(list_asn_status_user1)]
asn_user1_doc = set(asn_user1['ASN Number'])
asn_user2 = mis_asn2[mis_asn2['ASN Status'].isin(list_asn_status_user2)]
asn_user2_doc = set(asn_user2['ASN Number'])
print("asn mis common",len(asn_user2_doc))

#open-invoice filter
open_invoice_docket = set(open_invoice['Docket No.'])
status = ['Pending with SSC']
reason = ['WIP for header']
mis_status_open = old_mis[old_mis['Status of Invoice'].isin(status)]
mis_reason_open = mis_status_open[mis_status_open['Reason of Rejection/Hold'].isin(reason)]
mis_reason_docket = set(mis_reason_open['Docket No.'])
mis_open_docket = mis_reason_docket.intersection(open_invoice_docket)
mis_open_docket = list(mis_open_docket)
# mis_open = open_invoice[open_invoice['Docket No.'].isin(mis_open_docket)]
mis_open = mis_asn[mis_asn['Docket No.'].isin(mis_open_docket)]

#open-invoice validate
list_open_status_phd = ['Phd Rejects']
list_open_status_ssc =['Ssc validator']
open_status_phd = mis_open[mis_open['Status of Invoice'].isin(list_open_status_phd)]
open_phd_doc = set(open_status_phd['Docket No.'])
open_status_ssc = mis_open[mis_open['Status of Invoice'].isin(list_open_status_ssc)]
open_ssc_doc = set(open_status_ssc['Docket No.'])

#ssc combined mis with mis
ssc_hold = ['Hold_Processing','Hold_Validation','Rejection_Processing','Rejection_Validation']
ssc_pending = ['Pending for Processing','Ready for Validation']
ssc_validated = ['Validated']
ssc_status_hold = ssc_mis_combined[ssc_mis_combined['Status of Invoice'].isin(ssc_hold)]
ssc_status_hold_docket = set(ssc_status_hold['Docket No.'])
ssc_status_pending = ssc_mis_combined[ssc_mis_combined['Status of Invoice'].isin(ssc_pending)]
ssc_status_pending_docket = set(ssc_status_pending['Docket No.'])
ssc_status_validated = ssc_mis_combined[ssc_mis_combined['Status of Invoice'].isin(ssc_validated)]
ssc_status_val_docket = set(ssc_status_validated['Docket No.'])

#
#
master = pd.read_excel(r"C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\status_of_invoice_files\data (19).xlsx",sheet_name='Export')
# new_header = master.iloc[0]
# master = master[1:]
# master.columns = new_header
len_doc = len(master['Docket No.'])
print(len_doc)
for i in range(1,len_doc):
    #phd competency loop
    print("phd competency",master['Docket No.'][i])
    if master['Docket No.'][i] in set_hold:
        master.loc[i, 'Status of Invoice'] = "Hold at SSC"
        # master['Reason of Reason of Rejection/Hold'][i] ==
    elif master['Docket No.'][i] in set_otc:
        master.loc[i, 'Status of Invoice'] =  "Rejected to Partner"
        master.loc[i, 'Reason of Rejection/Hold'] =  "Rejected to Partner"
    elif master['Docket No.'][i] in set_scm:
        master.loc[i, 'Status of Invoice'] =  "Pending with SCM"
        master.loc[i, 'Reason of Rejection/Hold'] =  "Pending with GRN"
    elif master['Docket No.'][i] in set_user:
        master.loc[i, 'Status of Invoice'] =  "Pending with User"
        master.loc[i, 'Reason of Rejection/Hold'] =  "Pending with GRN"

print("phd competency time:",datetime.now())

for i in range(1, len_doc):
    #ssc_batch_loop
        print("ssc batch",master['Docket No.'][i])
        if master['Docket No.'][i] in ssc_batch_mis:
            master.loc[i, 'Status of Invoice'] =  "Pending with SSC"
            master.loc[i, 'Reason of Rejection/Hold'] =  "Ready for SSC"

print("ssc batch time:",datetime.now())
for i in range(1, len_doc):
    #eb_gstBatch loop
        print("eb gst batch",master['Docket No.'][i])
        if master['Docket No.'][i] in eb_gst_batch_mis:
            master.loc[i, 'Status of Invoice'] =  "Rejected to Partner"
            master.loc[i, 'Reason of Rejection/Hold'] =  "Belong to EB-GST/already to EB GST"

print("eb gst batch time:",datetime.now())
    # eb-gst-mis-loop
    # if master['Docket No.'][i] == eb_gst_mis:
    #     master.loc[i, 'Status of Invoice'] =  "Rejected to Partner"
    #     master.loc[i, 'Reason of Rejection/Hold'] =  "Belong to EB-GST/already to EB GST"
    # #i-expense loop
    #     if master['Docket No.'][i] == i_expense_mis:
    #         master.loc[i, 'Status of Invoice'] =  "Rejected to Partner"
    #         master.loc[i, 'Reason of Rejection/Hold'] =  "Belong to I expense"
for i in range(1, len_doc):
    #eastern mis
        print("eastern mis",master['Docket No.'][i])
        if master['Docket No.'][i] in eastern_mis_docket:
            mid_eastern = pd.DataFrame()
            check_list = []
            check_list.append(master['Docket No.'][i])
            print(check_list)
            mid_eastern = eastern_mis_columns[eastern_mis_columns['Docket No.'].isin(check_list)]
            # mid_eastern.to_excel('mid_eastern.xlsx')
            print(mid_eastern)
            mid_eastern.drop_duplicates()
            stat = mid_eastern['Status of Invoice'].to_list()
            print(stat[0])
            master.loc[i, 'Status of Invoice'] = stat[0]
            reason = mid_eastern['Reason of Rejection/Hold'].to_list()
            print(reason[0])
            if reason[0] == "Pre-merger ER data":
                master.loc[i, 'Reason of Rejection/Hold'] = "Migration Hold"
            remark = mid_eastern['Current Stage'].to_list()
            print(remark[0])
            master.loc[i, 'Current Stage'] = remark[0]
            # remark1 = list(remark.keys())[0]
            # print(remark1)
            # master.loc[i, 'Reason of Rejection/Hold'] = reason[0]
            # if master['Reason of Rejection/Hold'][i] == "Pre-merger ER data":
            #     master.loc[i, 'Reason of Rejection/Hold'] = "Migration Hold"
            #     print("changed reason")





print("eastern mis time:",datetime.now())
options_status = ['Due For Payment','Not Due for Payment','Validated']
option_add_remark = ['Circle_SSL']
eastern_mis_check = master[master['Status of Invoice'].isin(options_status)]
eastern_mis_check_2 = master[master['Current Stage'].isin(option_add_remark)]
mis_eastern_add_check = set(eastern_mis_check_2['Docket No.'])
#
# # eastern-mis-checks
for i in range(1,len(master)):
    print("eastern mis check 2", master['Docket No.'][i])
    if master['Docket No.'][i] in mis_eastern_add_check:
        master.loc[i, 'Status of Invoice'] =  'Pending with SSC'
        master.loc[i, 'Reason of Rejection/Hold'] =  'WIP for Header'
print("eastern mis checks:",datetime.now())
# #grn-done
for i in range(0,len(master)):
    print("grn done", master['Docket No.'][i])
    if master['Docket No.'][i] in grn_done_docket:
        master.loc[i, 'Status of Invoice'] =  'Pending with SSC'
        master.loc[i, 'Reason of Rejection/Hold'] =  'WIP for Header'
print("grn done time:",datetime.now())

#asn-mis
for i in range(0,len(master)):
     print("asn mis", master['Docket No.'][i])
     if master['Docket No.'][i] in asn_scm_doc:
         master.loc[i, 'Status of Invoice'] =  'Pending with SCM'
         master.loc[i, 'Reason of Rejection/Hold'] =  'Pending with GRN'
     elif master['Docket No.'][i] in asn_user1:
         master.loc[i, 'Status of Invoice'] =  'Pending with User'
         master.loc[i, 'Reason of Rejection/Hold'] =  'Pending for ASN Approval'
     elif master['Docket No.'][i] in asn_user2:
         master.loc[i, 'Status of Invoice'] =  'Pending with User'
         master.loc[i, 'Reason of Rejection/Hold'] =  'ASN Cancelled/Rejected'
print("asn time:",datetime.now())

#open invoice report
for i in range(0,len(master)):
     print("open invoice report", master['Docket No.'][i])
     if master['Docket No.'][i] in open_phd_doc:
         master.loc[i, 'Status of Invoice'] =  'Rejected to Partner'
         master.loc[i, 'Reason of Rejection/Hold'] ='Rejected to Partner'
     elif master['Docket No.'][i] in open_ssc_doc:
         master.loc[i, 'Status of Invoice'] =  'Pending with SSC'
         master.loc[i, 'Reason of Rejection/Hold'] = 'Ready for SSC'
print("open invoice time :",datetime.now())
# ssc combined update
# for i in range(0,len(master)):
#      print("ssc combined", master['Docket No.'][i])
#      if master['Docket No.'][i] in ssc_status_hold_docket:
#          master.loc[i, 'Status of Invoice'] =  'Hold at SSC'
#          master.loc[i, 'Reason of Rejection/Hold'] =  'Balance Confirmation Hold'
#      elif master['Docket No.'][i] in ssc_status_pending_docket:
#          master.loc[i, 'Status of Invoice'] =  'Pending with SSC'
#          master.loc[i, 'Reason of Rejection/Hold'] =  'WIP for Header'
     # elif master['Docket No.'][i] in ssc_status_val_docket:
     #     master.loc[i, 'Status of Invoice'] =  'Validated'
for i in range(0,len(master)):
    print("ssc combined", master['Docket No.'][i])
    if master['Docket No.'][i] in ssc_status_hold_docket:
         master.loc[i, 'Status of Invoice'] =  'Hold at SSC'
         master.loc[i, 'Reason of Rejection/Hold'] =  'Validated'
for i in range(0,len(master)):
     print("ssc combined", master['Docket No.'][i])
     if master['Docket No.'][i] in ssc_status_hold_docket:
         master.loc[i, 'Status of Invoice'] =  'Hold at SSC'
         ssc_reason = pd.DataFrame()
         check_list2 = []
         check_list2.append(master['Docket No.'][i])
         print(check_list2)
         ssc_reason = ssc_mis_combined[ssc_mis_combined['Docket No.'].isin(check_list2)]
         # mid_eastern.to_excel('mid_eastern.xlsx')
         print(ssc_mis_combined)
         ssc_mis_combined.drop_duplicates()
         reason_ssc = mid_eastern['Reason of Rejection/Hold'].to_list()
         print(reason_ssc[0])
         master.loc[i, 'Reason of Rejection/Hold'] = reason_ssc[0]
     elif master['Docket No.'][i] in ssc_status_pending_docket:
         master.loc[i, 'Status of Invoice'] =  'Pending with SSC'
         master.loc[i, 'Reason of Rejection/Hold'] =  'WIP for Header'
     elif master['Docket No.'][i] in ssc_status_val_docket:
         master.loc[i, 'Status of Invoice'] =  'Validated'
         master.loc[i, 'Reason of Rejection/Hold'] =  'Validated'

for i in range(0,len(master)):
     print("ssc combined balance confirmation", master['Docket No.'][i])
     if master['Reason of Rejection/Hold'][i] == "Balance confirmation":
         master.loc[i, 'Status of Invoice'] =  "Balance confirmation Hold"


print("ssc combined time:",datetime.now())
# ############################################################################################
# #try paid above

# paid_inward = set(paid['Inward No'])
# paid_mis_docket = mis_docket.intersection(paid_inward)
# for i in range(0,len(master)):
#     print("paid due date erp", master['Docket No.'][i])
#     if master['Docket No.'][i] in paid_mis_docket:
#         if master['Due Date ERP'][i] <= date_today:
#             master.loc[i, 'Status of Invoice'] =  'Due for Payment'
#             master.loc[i, 'Reason of Rejection/Hold'] =  'Due for Payment'
#         elif master['Due Date ERP'][i] > date_today:
#             master.loc[i, 'Status of Invoice'] =  'Not Due for Payment'
#             master.loc[i, 'Reason of Rejection/Hold'] =  'Not Due for Payment'
# for i in range(0, len(master)):
#     print("paid payment hold", master['Docket No.'][i])
#     if master['Docket No.'][i] in paid_mis_docket:
#         if master['Payment Hold Flag'][i] == "Yes":
#             master.loc[i, 'Status of Invoice'] =  'Hold by Backend/Circle Hold'
# print("car paid check :",datetime.now())
#
# paid_mis_docket = list(paid_mis_docket)
# mis_paid = master[master['Docket No.'].isin(paid_mis_docket)]
# status_mis = ['Due for Payment','Not Due for Payment','Hold by Backend/Circle Hold']
# mis_paid2 = mis_paid[mis_paid['Status of Invoice'].isin(status_mis)]
# mis_paid3 = mis_paid2[mis_paid2['Creditor Status of Invoice'].isin(creditor_status)]
# print(mis_paid3)
# paid_clearance = set(mis_paid3['Docket No.'])

# for i in range(0, len(master)):
#     print("paid payment hold paid under clearance", master['Docket No.'][i])
#     if master['Docket No.'][i] in paid_clearance:
#             master.loc[i, 'Status of Invoice'] =  'Paid Under Clearance'
#
# print("paid under clearance updated :",datetime.now())

# # now2 = datetime.now()
# # print("start time:",now1)
# # print("end time:",now2)
# # print("all status updateddddd!!!!!")


#try end
#################################################################################################
car_docket = set(car['Document ID'])
car_oracle = set(car['Oracle Invoice ID'])
car_val = car_docket.union(car_oracle)
mis_car_docket = mis_docket.intersection(car_val)
print("mis_car_docket len",len(mis_car_docket))
mis_car_docket = list(mis_car_docket)
car_common_mis = car[car['Document ID'].isin(mis_car_docket)]
car_common_mis2 = car[car['Oracle Invoice ID'].isin(mis_car_docket)]
car_mis = pd.concat([car_common_mis, car_common_mis2]).drop_duplicates()
car_mis_document = set(car_mis['Document ID'])
car_mis_oracle = set(car_mis['Oracle Invoice ID'])
car_mis_docket = car_mis_document.union(car_mis_oracle)
car_mis_docket = list(car_mis_docket)

creditor_status = ['Not Found']
creditor_mis = old_mis[old_mis['Docket No.'].isin(car_mis_docket)]
car_mis = creditor_mis = old_mis[old_mis['Creditor Status of Invoice'].isin(creditor_status)]
car_invoice = set(car['Invoice Number'])
mis_car_invoice = mis_invoice_num.intersection(car_invoice)
mis_car_invoice = list(mis_car_invoice)
car_mis2 = car_mis[car_mis['Invoice No.'].isin(mis_car_invoice)]
creditor_mis_invoice = car_mis2[car_mis2['Creditor Status of Invoice'].isin(creditor_status)]
car['combo'] = str(car['Circle'])+ str(car['Vendor Code']) + str(car['Invoice Number'])
car_combo_docket = set(car['combo'])
mis_car_combo = mis_docket.intersection(car_combo_docket)
mis_car_combo = list(mis_car_combo)
print(mis_car_combo)
creditor_mis_combo = car_mis2[car_mis2['Docket No.'].isin(mis_car_combo)]

status_car1 = ['Validated']
car_status_validated = creditor_mis_combo[creditor_mis_invoice['Docket No.'].isin(status_car1)]
car_status_validated_docket = set(car_status_validated['Docket No.'])
print("car validated step 1")

#car step one
for i in range(0,len(master)):
    print("car loop",master['Docket No.'][i])
    if master['Docket No.'][i] in car_status_validated_docket:
        master.loc[i, 'Status of Invoice'] =  'Validated'
        master.loc[i, 'Reason of Rejection/Hold'] =  'Validated'
print("car step one time:",datetime.now())

car_status_creditor = ['Needs Revalidation','Unvalidated']
filter1 = creditor_mis[creditor_mis['Creditor Status of Invoice'].isin(car_status_creditor)]
status_of_invoice = ['Due for Payment','Not Due for Payment','Hold by Backend/Circle Hold','Paid Under Clearance']
filter2 = filter1[filter1['Status of Invoice'].isin(status_of_invoice)]
car_status_docket2 = set(filter2['Docket No.'])
print("car validation step 2")
car_status_cr = ['Validated']
stat_of_inv = ['Due for Payment','Not Due for Payment','Validated']
car3 = car[car['Status Of Invoice'].isin(car_status_cr)]
master3 = master[master['Status of Invoice'].isin(stat_of_inv)]
car3set = set(car3['Document ID'])
car3set_oracle = set(car3['Oracle Invoice ID'])
car3set_union = car3set.union(car3set_oracle)
car_due_not_due = mis_docket.intersection(car3set_union)


for i in range(0,len(master)):
    print("car loop 2",master['Docket No.'][i])
    if master['Docket No.'][i] in car_status_docket2:
        master.loc[i, 'Status of Invoice'] =  'Pending with SSC'
        master.loc[i, 'Reason of Rejection/Hold'] = 'WIP for Processing'
print("car step two time:",datetime.now())
# date_today = date.today()
master.to_excel('status_of_invoice_updatedwithcar.xlsx')

date_today = date(2022, 7, 5)
print(date_today)
#try
for i in range(0, len(master)):
    print("car loop due payment hold", master['Docket No.'][i])
    if master['Docket No.'][i] in car_mis_docket:
        if master['Payment Hold Flag'][i] == 'Yes':
            master.loc[i, 'Status of Invoice'] =  'Hold by Backend/Circle Hold'

master_status = master[master['Status of Invoice'] == "Hold at SSC"]
master_reject = master_status[master_status['Reason of Rejection/Hold'] == "workflow approval pending with user"]
master_car_wf = mis_docket.intersection(car_mis_docket)

for i in range(0,len(master)):
    print("car loop due date erp", master['Docket No.'][i])
    if master['Docket No.'][i] in car_due_not_due:
        if master['Due Date ERP'][i] <= date_today:
            master.loc[i, 'Status of Invoice'] =  'Due for Payment'
            master.loc[i, 'Reason of Rejection/Hold'] =  'Due for Payment'
        elif master['Due Date ERP'][i] > date_today:
            master.loc[i, 'Status of Invoice'] =  'Not Due for Payment'
            master.loc[i, 'Reason of Rejection/Hold'] =  'Not Due for Payment'
print("car due date time:",datetime.now())
for i in range(0, len(master)):
    print("car loop due payment hold", master['Docket No.'][i])
    if master['Docket No.'][i] in car_mis_docket:
        if master['Payment Hold Flag'][i] == 'Yes':
            master.loc[i, 'Status of Invoice'] =  'Hold by Backend/Circle Hold'
print("car payment hold flag time:",datetime.now())
master_status = master[master['Status of Invoice'] == "Hold at SSC"]
master_reject = master_status[master_status['Reason of Rejection/Hold'] == "Workflow approval pending with user"]
master_car_wf = mis_docket.intersection(car_mis_docket)
master_car_wf = list(master_car_wf)
print(master_car_wf)
for i in range(0, len(master)):
    print("car loop due wf status", master['Docket No.'][i])
    if master['Docket No.'][i] in master_car_wf:
        # if car['WF Status'] == "WFAPPROVED":
            mid_car = pd.DataFrame()
            check_list1 = []
            check_list1.append(str(master['Docket No.'][i]))
            print(check_list1)
            mid_car = car[car['Document ID'].isin(check_list1)]
            mid_car_or = car[car['Oracle Invoice ID'].isin(check_list1)]
            mid_car_comb = pd.concat([mid_car, mid_car_or]).drop_duplicates()
            l = mid_car_comb['WF Status'].values.tolist()
            l_check = ['WFAPPROVED']
            # status = l[0]
            # print(status)
            print(l)
            if l == l_check:
                master.loc[i, 'Status of Invoice'] =  'Pending with SSC'
                master.loc[i, 'Reason of Rejection/Hold'] =  'workflow approval'
            # if mid_car_or['WF Status'].iloc[0] == "WFAPPROVED":
            # if m[0] == "WFAPPROVED":
            #     master.loc[i, 'Status of Invoice'] = 'Pending with SSC'
            #     master.loc[i, 'Reason of Rejection/Hold'] = 'workflow approval'

print("car done")
print("car completion time:",datetime.now())
#####################################################################################
paid_inward = set(paid['Inward No'])
paid_mis_docket = mis_docket.intersection(paid_inward)
for i in range(0,len(master)):
    print("paid due date erp", master['Docket No.'][i])
    if master['Docket No.'][i] in paid_mis_docket:
        if master['Due Date ERP'][i] <= date_today:
            master.loc[i, 'Status of Invoice'] =  'Due for Payment'
            master.loc[i, 'Reason of Rejection/Hold'] =  'Due for Payment'
        elif master['Due Date ERP'][i] > date_today:
            master.loc[i, 'Status of Invoice'] =  'Not Due for Payment'
            master.loc[i, 'Reason of Rejection/Hold'] =  'Not Due for Payment'
for i in range(0, len(master)):
    print("paid payment hold", master['Docket No.'][i])
    if master['Docket No.'][i] in paid_mis_docket:
        if master['Payment Hold Flag'][i] == "Yes":
            master.loc[i, 'Status of Invoice'] =  'Hold by Backend/Circle Hold'
print("car paid check :",datetime.now())

paid_mis_docket = list(paid_mis_docket)
mis_paid = master[master['Docket No.'].isin(paid_mis_docket)]
status_mis = ['Due for Payment','Not Due for Payment','Hold by Backend/Circle Hold']
mis_paid2 = mis_paid[mis_paid['Status of Invoice'].isin(status_mis)]
mis_paid3 = mis_paid2[mis_paid2['Creditor Status of Invoice'].isin(creditor_status)]
print(mis_paid3)
paid_clearance = set(mis_paid3['Docket No.'])

for i in range(0, len(master)):
    print("paid payment hold paid under clearance", master['Docket No.'][i])
    if master['Docket No.'][i] in paid_clearance:
            master.loc[i, 'Status of Invoice'] =  'Paid Under Clearance'

print("paid under clearance updated :",datetime.now())

now2 = datetime.now()
print("start time:",now1)
print("end time:",now2)
print("all status updateddddd!!!!!")
#####################################################################################################
master.to_excel(r"C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\MIS DATA\OUTPUTS\status output\status_of_invoice_updated10.xlsx")
# sys.stdout.close()


print("hogya biro dekhlo fatafat")

# SENDER_EMAIL = "kshiti.sinha@myndsol.com"
# APP_PASSWORD = "1@Million"
# date_today = datetime.date
#
# excel_file=r"C:\Users\kshiti.sinha\Desktop\projects\MIS TRACKER\MIS DATA\OUTPUTS\status output\status_of_invoice_updated6.xlsx"
# subject = "MIS DATA"
# recipient_email = "analytics@myndsol.com"
#
# def send_mail_with_excel():
#     msg = EmailMessage()
#     msg['Subject'] = 'MIS DATA'
#     msg['From'] = SENDER_EMAIL
#     msg['To'] = "analytics@myndsol.com"
#
#
#     with open(excel_file, 'rb') as f:
#         file_data = f.read()
#     msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename='MIS DATA')
#
#     with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
#         smtp.login(SENDER_EMAIL, APP_PASSWORD)
#         smtp.send_message(msg)
#
# send_mail_with_excel()

#send email





