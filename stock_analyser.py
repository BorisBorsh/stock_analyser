import loadfile
import os
from operations_on_companies_list import *
from openpyxl import load_workbook
from openpyxl.styles import GradientFill
from openpyxl.styles import Font
from evaluation_properties import StandardEvaluationModel

excel_file_path = r'C:\\Champions.xlsx'
loadFile_permission = False

wb = load_workbook(filename=excel_file_path, data_only=True)
ws = wb['Champions']

stnd_model = StandardEvaluationModel()
prelim_champs_list = []

green_grad_bg_fill = GradientFill(stop=("0fe00b", "FFFFFF"))


prelim_champs_list = get_fundamental_analysis_on_champs_list(ws, stnd_model)

champ_list = prelim_champs_list
        #ws['A' + str(i)].font = Font(b=True, color='0fe00b')
        #ws['A' + str(i)].fill = green_grad_bg_fill
        # FFFF0000 red font color


ws_hist = wb['Historical']
company_hist_list_start_indx = first_company_in_list_cell(ws_hist)
company_hist_list_end_indx = last_company_in_hist_list_cell(ws_hist)

#Percentage Increase by Year Analysis (5 years row)
result_array = []
for company in range(0, len(champ_list)):
    i = find_company_in_list(champ_list[company]['company_name'], ws_hist, company_hist_list_start_indx, company_hist_list_end_indx)

    if (
        ws_hist['AA' + str(i)].value >= stnd_model.percentage_increase_by_year
        and ws_hist['AB' + str(i)].value >= stnd_model.percentage_increase_by_year
        and ws_hist['AC' + str(i)].value >= stnd_model.percentage_increase_by_year
        and ws_hist['AD' + str(i)].value >= stnd_model.percentage_increase_by_year
        and ws_hist['AE' + str(i)].value >= stnd_model.percentage_increase_by_year
    ):
        result_array.append(champ_list[company])
        ws_hist['AA' + str(i)].fill = green_grad_bg_fill
        ws_hist['AB' + str(i)].fill = green_grad_bg_fill
        ws_hist['AC' + str(i)].fill = green_grad_bg_fill
        ws_hist['AD' + str(i)].fill = green_grad_bg_fill
        ws_hist['AE' + str(i)].fill = green_grad_bg_fill

champ_list = result_array


#Year by year dividend growth
result_array = []
for company in range(0, len(champ_list)):
    i = find_company_in_list(champ_list[company]['company_name'], ws_hist, company_hist_list_start_indx, company_hist_list_end_indx)

    if (
        ws_hist['C' + str(i)].font.color is None
        and ws_hist['D' + str(i)].font.color is None
        and ws_hist['E' + str(i)].font.color is None
        and ws_hist['F' + str(i)].font.color is None
        and ws_hist['G' + str(i)].font.color is None
        and ws_hist['H' + str(i)].font.color is None
        and ws_hist['I' + str(i)].font.color is None
        and ws_hist['K' + str(i)].font.color is None
        and ws_hist['L' + str(i)].font.color is None
        and ws_hist['N' + str(i)].font.color is None
        and ws_hist['O' + str(i)].font.color is None
        and ws_hist['P' + str(i)].font.color is None
        and ws_hist['Q' + str(i)].font.color is None
        and ws_hist['R' + str(i)].font.color is None
        and ws_hist['S' + str(i)].font.color is None
        and ws_hist['T' + str(i)].font.color is None
        and ws_hist['U' + str(i)].font.color is None
        and ws_hist['V' + str(i)].font.color is None
        and ws_hist['W' + str(i)].font.color is None
        and ws_hist['X' + str(i)].font.color is None
        and ws_hist['Y' + str(i)].font.color is None
    ):
        result_array.append(champ_list[company])
        ws_hist['A' + str(i)].fill = green_grad_bg_fill
        ws_hist['C' + str(i)].fill = green_grad_bg_fill
        ws_hist['D' + str(i)].fill = green_grad_bg_fill
        ws_hist['E' + str(i)].fill = green_grad_bg_fill
        ws_hist['F' + str(i)].fill = green_grad_bg_fill
        ws_hist['G' + str(i)].fill = green_grad_bg_fill
        ws_hist['H' + str(i)].fill = green_grad_bg_fill
        ws_hist['I' + str(i)].fill = green_grad_bg_fill
        ws_hist['J' + str(i)].fill = green_grad_bg_fill
        ws_hist['K' + str(i)].fill = green_grad_bg_fill
        ws_hist['L' + str(i)].fill = green_grad_bg_fill
        ws_hist['M' + str(i)].fill = green_grad_bg_fill
        ws_hist['N' + str(i)].fill = green_grad_bg_fill
        ws_hist['O' + str(i)].fill = green_grad_bg_fill
        ws_hist['P' + str(i)].fill = green_grad_bg_fill
        ws_hist['Q' + str(i)].fill = green_grad_bg_fill
        ws_hist['R' + str(i)].fill = green_grad_bg_fill
        ws_hist['S' + str(i)].fill = green_grad_bg_fill
        ws_hist['T' + str(i)].fill = green_grad_bg_fill
        ws_hist['U' + str(i)].fill = green_grad_bg_fill
        ws_hist['V' + str(i)].fill = green_grad_bg_fill
        ws_hist['W' + str(i)].fill = green_grad_bg_fill
        ws_hist['X' + str(i)].fill = green_grad_bg_fill
        ws_hist['Y' + str(i)].fill = green_grad_bg_fill

champ_list = result_array

print("Saving data to a C:\Result.xlsx file")

wb.save('C:\\Result.xlsx')
print("DONE!")