import loadfile
import os
from operations_on_companies_list import *
from openpyxl import load_workbook
from openpyxl.styles import GradientFill
from openpyxl.styles import Font
from evaluation_properties import StandardEvaluationModel

excel_file_path = r'C:\\Champions.xlsx'
loadFile_permission = False


wb = load_workbook(filename='C:\Champions.xlsx', data_only=True)

stnd_model = StandardEvaluationModel()

ws = wb['Champions']
ws_hist = wb['Historical']
champ_list = []

green_grad_bg_fill = GradientFill(stop=("0fe00b", "FFFFFF"))


company_list_start_indx = first_company_in_list_cell(ws)
company_list_end_indx = last_company_in_list_cell(ws)
print("company_list_start_indx = " + str(company_list_start_indx))
print("company_list_start_indx = " + str(company_list_end_indx))

print("Analysing data")

for i in range(company_list_start_indx, company_list_end_indx):

    if (    #Div Years
            ws['E' + str(i)].value >= stnd_model.div_years
            #Overall AVG divs
            and float(ws['J' + str(i)].value) >= stnd_model.avg_divs_overall
            #MR last dividends inc%
            and float(ws['R' + str(i)].value) >= stnd_model.mr_last_div_incr
            #EPS a part of profit to dividends
            and ws['Z' + str(i)].value != 'n/a'
            and float(ws['Z' + str(i)].value) < stnd_model.eps
            #P/E AVG
            and float(ws['AA' + str(i)].value) <= stnd_model.pe_avg
            #MktCap, $Mil
            and float(ws['AL' + str(i)].value) >= stnd_model.cap_mil_dollrs
            #Est. div in 5 years Payback, %
            and float(ws['AX' + str(i)].value) >= stnd_model.est_div_paybacks_5years_predicted
    ):
        champion = dict()
        champion['company_name'] = ws['A' + str(i)].value
        champion['div_years_row'] = ws['E' + str(i)].value
        champion['dividends_avg'] = to_fixed(ws['J' + str(i)].value,2)
        champion['MR%'] = to_fixed(ws['R' + str(i)].value, 2)
        champion['EPS'] = to_fixed(ws['Z' + str(i)].value, 2)
        champion['AVG_PE'] = to_fixed(ws['AA' + str(i)].value, 2)
        champion['capitalization_mil$'] = to_fixed(ws['AL' + str(i)].value, 2)
        champion['est_divs'] = to_fixed(ws['AX' + str(i)].value, 2)
        #ws['A' + str(i)].font = Font(b=True, color='0fe00b')

        ws['A' + str(i)].fill = green_grad_bg_fill
        #print(ws['A' + str(i)].value)
        #FFFF0000 red font color
        #print(ws_hist['I15'].font.color.rgb)
        champ_list.append(champion)


#Looking for 5% divs in row for previously selected champions
company_hist_list_start_indx = first_company_in_list_cell(ws_hist)
company_hist_list_end_indx = last_company_in_hist_list_cell(ws_hist)
#print("company_hist_list_start_indx = " + str(company_hist_list_start_indx))
#print("company_hist_list_end_indx = " + str(company_hist_list_end_indx))

for company in range(0, len(champ_list)):
    print('companyName ->', champ_list[company]['company_name'])
    #print(champ_list[company])
    print("position is ", find_company_in_list(champ_list[company]['company_name'], ws_hist, company_hist_list_start_indx, company_hist_list_end_indx))

print("Saving data to a C:\Result.xlsx file")

wb.save('C:\\Result.xlsx')



