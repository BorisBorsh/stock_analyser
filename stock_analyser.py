import loadfile
import os
from operations_on_companies_list import *
from analyse_and_highlight_data import get_champ_list_after_fundamental_analysis, \
                                color_fundamental_parameters_of_companies_in_list, \
                                get_champ_list_after_5years_dividends_increase_in_row_analysis, \
                                color_params_of_champ_list_5years_dividends_increase_in_row
from openpyxl import load_workbook
from openpyxl.styles import GradientFill
from openpyxl.styles import Font
from evaluation_properties import StandardEvaluationModel

excel_file_path = r'C:\\Champions.xlsx'
download_file_permission = False

wb = load_workbook(filename=excel_file_path, data_only=True)
ws = wb['Champions']
ws_hist = wb['Historical']

stnd_model = StandardEvaluationModel()


#Colored fill for cells
green_grad_bg_fill = GradientFill(stop=("0fe00b", "FFFFFF"))
blue_grad_bg_fill = GradientFill(stop=("00e7fe", "FFFFFF"))

#Get a preliminary list of champions according to fundamental  analysis of companies parameters
prelim_champs_list = get_champ_list_after_fundamental_analysis(ws, stnd_model)
#Fill with color all the cells that passed fundamental requirements
color_fundamental_parameters_of_companies_in_list(prelim_champs_list, ws, blue_grad_bg_fill)

#Check if preliminary list of companies are met conditions of 5 years in row of dividend increase
post_prelim_champs_list = get_champ_list_after_5years_dividends_increase_in_row_analysis(prelim_champs_list,
                                                                                         stnd_model, ws_hist)
#Color fill all the cells that met conditions of 5 years in row of dividend increase
get_champ_list_after_5years_dividends_increase_in_row_analysis(post_prelim_champs_list, ws_hist, green_grad_bg_fill)
#
color_params_of_champ_list_5years_dividends_increase_in_row(post_prelim_champs_list, ws_hist, blue_grad_bg_fill)


champ_list = post_prelim_champs_list

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
