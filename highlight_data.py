from operations_on_companies_list import first_company_in_list_cell, last_company_in_list_cell \
    , find_company_in_list, last_company_in_hist_list_cell


def color_fundamental_parameters_of_companies_in_list(companies_list, champ_worksheet, color_fill):
    """Highlight company name and all fundamental parameters of company,
        list of companies in parameters of the function"""

    ws = champ_worksheet
    company_list_start_indx = first_company_in_list_cell(ws)
    company_list_end_indx = last_company_in_list_cell(ws)

    for company in range(0, len(companies_list)):
        i = find_company_in_list(companies_list[company]['company_name'], ws, company_list_start_indx,
                                 company_list_end_indx)
        # Fill cells
        # Highlight company name
        ws['A' + str(i)].fill = color_fill
        # Div Years
        ws['E' + str(i)].fill = color_fill
        # Overall AVG divs
        ws['J' + str(i)].fill = color_fill
        # MR last dividends inc%
        ws['R' + str(i)].fill = color_fill
        # EPS a part of profit to dividends
        ws['Z' + str(i)].fill = color_fill
        # P/E AVG
        ws['AA' + str(i)].fill = color_fill
        # MktCap, $Mil
        ws['AL' + str(i)].fill = color_fill
        # Est. div in 5 years Payback, %
        ws['AX' + str(i)].fill = color_fill
        # ws['A' + str(i)].font = Font(b=True, color='0fe00b')


def color_params_of_champ_list_5years_dividends_increase_in_row(input_champ_list, hist_worksheet, color_fill):
    """Color all params that were analysed during analisis of
        dividends magnification for 5 years in row"""

    ws_hist = hist_worksheet
    company_hist_list_start_indx = first_company_in_list_cell(ws_hist)
    company_hist_list_end_indx = last_company_in_hist_list_cell(ws_hist)

    for company in range(0, len(input_champ_list)):
        i = find_company_in_list(input_champ_list[company]['company_name'], ws_hist, company_hist_list_start_indx,
                                 company_hist_list_end_indx)

        ws_hist['AA' + str(i)].fill = color_fill
        ws_hist['AB' + str(i)].fill = color_fill
        ws_hist['AC' + str(i)].fill = color_fill
        ws_hist['AD' + str(i)].fill = color_fill
        ws_hist['AE' + str(i)].fill = color_fill


def color_champ_list_after_year_by_year_div_growth_analysis(input_champ_list, hist_worksheet, color_fill):
    """Color all params(dividends growth, %) that were analysed during year by year
        dividends growth is companies history"""

    ws_hist = hist_worksheet
    company_hist_list_start_indx = first_company_in_list_cell(ws_hist)
    company_hist_list_end_indx = last_company_in_hist_list_cell(ws_hist)

    for company in range(0, len(input_champ_list)):
        i = find_company_in_list(input_champ_list[company]['company_name'], ws_hist, company_hist_list_start_indx,
                                 company_hist_list_end_indx)

        ws_hist['A' + str(i)].fill = color_fill
        ws_hist['C' + str(i)].fill = color_fill
        ws_hist['D' + str(i)].fill = color_fill
        ws_hist['E' + str(i)].fill = color_fill
        ws_hist['F' + str(i)].fill = color_fill
        ws_hist['G' + str(i)].fill = color_fill
        ws_hist['H' + str(i)].fill = color_fill
        ws_hist['I' + str(i)].fill = color_fill
        ws_hist['J' + str(i)].fill = color_fill
        ws_hist['K' + str(i)].fill = color_fill
        ws_hist['L' + str(i)].fill = color_fill
        ws_hist['M' + str(i)].fill = color_fill
        ws_hist['N' + str(i)].fill = color_fill
        ws_hist['O' + str(i)].fill = color_fill
        ws_hist['P' + str(i)].fill = color_fill
        ws_hist['Q' + str(i)].fill = color_fill
        ws_hist['R' + str(i)].fill = color_fill
        ws_hist['S' + str(i)].fill = color_fill
        ws_hist['T' + str(i)].fill = color_fill
        ws_hist['U' + str(i)].fill = color_fill
        ws_hist['Y' + str(i)].fill = color_fill
        ws_hist['V' + str(i)].fill = color_fill
        ws_hist['W' + str(i)].fill = color_fill
        ws_hist['X' + str(i)].fill = color_fill
