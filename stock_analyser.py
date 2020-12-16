import os
from download_data_file import download_excel_data_file
from analyse_data import get_champ_list_after_fundamental_analysis, \
                         get_champ_list_after_5years_dividends_increase_in_row_analysis, \
                         get_final_champ_list_after_year_by_year_div_growth_analysis, \
                         get_final_champ_list_after_year_by_year_div_growth_analysis_contenders
from highlight_data import color_fundamental_parameters_of_companies_in_list, \
                           color_params_of_champ_list_5years_dividends_increase_in_row, \
                           color_champ_list_after_year_by_year_div_growth_analysis
from openpyxl import load_workbook
from evaluation_properties import StandardEvaluationModel


if __name__ == "__main__":

    print("Starting stock analyser")

    excel_data_file_path = r'C:\\Champions.xlsx'
    analyzed_excel_data_file_path = 'C:\\Result.xlsx'
    download_data_file_permission = True

    if download_data_file_permission:
        download_excel_data_file(excel_data_file_path)

    company_types = ['Champions', 'Contenders']
    wb = load_workbook(filename=excel_data_file_path, data_only=True)
    ws_hist = wb['Historical']

    for company_type in company_types:
        ws = wb[company_type]

        stnd_model = StandardEvaluationModel()
        stnd_model.div_years_adjustment(company_type)

        #Colored fill for cells
        final_grad_bg_fill, preliminary_grad_bg_fill = stnd_model.color_adjustment(company_type)

        print("Analysing ", company_type, " list")
        #PRELIM CHAMPIONS LOGIC
        #Get a preliminary list of champions according to fundamental  analysis of companies parameters
        prelim_champs_list = get_champ_list_after_fundamental_analysis(ws, stnd_model)
        #Fill with color all the cells that passed fundamental requirements
        color_fundamental_parameters_of_companies_in_list(prelim_champs_list, ws, preliminary_grad_bg_fill)
        #Color fill prelim chaps that could met conditions of 5 years in row of dividend increase
        color_champ_list_after_year_by_year_div_growth_analysis(prelim_champs_list, ws_hist, preliminary_grad_bg_fill)
        #Color fill prelim chaps that could met conditions of year by year div growth
        color_params_of_champ_list_5years_dividends_increase_in_row(prelim_champs_list, ws_hist, preliminary_grad_bg_fill)

        #POST_PRELIM CHAMPIONS LOGIC (5 YEARS DIV GROWTH)
        #Check if preliminary list of companies are met conditions of 5 years in row of dividend increase
        post_prelim_champs_list = get_champ_list_after_5years_dividends_increase_in_row_analysis(prelim_champs_list,
                                                                                                 stnd_model, ws_hist)
        #Color fill all the cells that met conditions of 5 years in row of dividend increase
        color_params_of_champ_list_5years_dividends_increase_in_row(post_prelim_champs_list, ws_hist, preliminary_grad_bg_fill)

        #FINAL CHAMPIONS LOGIC (YEAR BY YEAR DIV GROWTH)
        #Get a final list of champions according analysis of year by year div growth
        if company_type == 'Champions':
            final_champ_list = get_final_champ_list_after_year_by_year_div_growth_analysis(post_prelim_champs_list,
                                                                                           ws_hist)
        else:
            final_champ_list = get_final_champ_list_after_year_by_year_div_growth_analysis_contenders(post_prelim_champs_list,
                                                                                           ws_hist)
        #Color parameters of year by year div growth with green fill
        color_champ_list_after_year_by_year_div_growth_analysis(final_champ_list, ws_hist, final_grad_bg_fill)
        #Color parameters of year by year div growth with green fill (for final list)
        color_params_of_champ_list_5years_dividends_increase_in_row(final_champ_list, ws_hist, final_grad_bg_fill)
        #Color all the cells of fundamental parameters fo final champ list
        color_fundamental_parameters_of_companies_in_list(final_champ_list, ws, final_grad_bg_fill)

    print("Saving data to ", analyzed_excel_data_file_path)

    wb.save(analyzed_excel_data_file_path)
    print("DONE!")

    os.startfile(analyzed_excel_data_file_path)
