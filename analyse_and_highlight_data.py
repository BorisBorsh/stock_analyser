from operations_on_companies_list import first_company_in_list_cell, last_company_in_list_cell, \
    to_fixed, find_company_in_list, last_company_in_hist_list_cell


def get_champ_list_after_fundamental_analysis(champ_worksheet, eval_model):
    """Analysing preliminary list of potential champs in stock list using fundamental analysis parameters"""

    ws = champ_worksheet
    resulted_champ_list = []
    company_list_start_indx = first_company_in_list_cell(ws)
    company_list_end_indx = last_company_in_list_cell(ws)

    for i in range(company_list_start_indx, company_list_end_indx):

        if (  # Div Years
                ws['E' + str(i)].value >= eval_model.div_years
                # Overall AVG divs
                and float(ws['J' + str(i)].value) >= eval_model.avg_divs_overall
                # MR last dividends inc%
                and float(ws['R' + str(i)].value) >= eval_model.mr_last_div_incr
                # EPS a part of profit to dividends
                and ws['Z' + str(i)].value != 'n/a'
                and float(ws['Z' + str(i)].value) < eval_model.eps
                # P/E AVG
                and float(ws['AA' + str(i)].value) <= eval_model.pe_avg
                # MktCap, $Mil
                and float(ws['AL' + str(i)].value) >= eval_model.cap_mil_dollrs
                # Est. div in 5 years Payback, %
                and float(ws['AX' + str(i)].value) >= eval_model.est_div_paybacks_5years_predicted
        ):
            champion = dict()
            champion['company_name'] = ws['A' + str(i)].value
            champion['div_years_row'] = ws['E' + str(i)].value
            champion['dividends_avg'] = to_fixed(ws['J' + str(i)].value, 2)
            champion['MR%'] = to_fixed(ws['R' + str(i)].value, 2)
            champion['EPS'] = to_fixed(ws['Z' + str(i)].value, 2)
            champion['AVG_PE'] = to_fixed(ws['AA' + str(i)].value, 2)
            champion['capitalization_mil$'] = to_fixed(ws['AL' + str(i)].value, 2)
            champion['est_divs'] = to_fixed(ws['AX' + str(i)].value, 2)
            resulted_champ_list.append(champion)

    return resulted_champ_list


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


def get_champ_list_after_5years_dividends_increase_in_row_analysis(input_champ_list, eval_model, hist_worksheet):
    """Analysing dividends magnification for 5 years in row"""

    ws_hist = hist_worksheet
    company_hist_list_start_indx = first_company_in_list_cell(ws_hist)
    company_hist_list_end_indx = last_company_in_hist_list_cell(ws_hist)

    resulted_champ_list_div_row = []
    for company in range(0, len(input_champ_list)):
        i = find_company_in_list(input_champ_list[company]['company_name'], ws_hist, company_hist_list_start_indx,
                                 company_hist_list_end_indx)

        if (
                ws_hist['AA' + str(i)].value >= eval_model.percentage_increase_by_year
                and ws_hist['AB' + str(i)].value >= eval_model.percentage_increase_by_year
                and ws_hist['AC' + str(i)].value >= eval_model.percentage_increase_by_year
                and ws_hist['AD' + str(i)].value >= eval_model.percentage_increase_by_year
                and ws_hist['AE' + str(i)].value >= eval_model.percentage_increase_by_year
        ):
            resulted_champ_list_div_row.append(input_champ_list[company])

    return resulted_champ_list_div_row


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


# Year by year dividend growth
def get_final_champ_list_after_year_by_year_div_growth_analysis(input_champ_list, hist_worksheet):
    """"""

    ws_hist = hist_worksheet
    company_hist_list_start_indx = first_company_in_list_cell(ws_hist)
    company_hist_list_end_indx = last_company_in_hist_list_cell(ws_hist)
    champ_array_year_by_year_div_growth = []
    for company in range(0, len(input_champ_list)):
        i = find_company_in_list(input_champ_list[company]['company_name'], ws_hist, company_hist_list_start_indx,
                                 company_hist_list_end_indx)

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
            champ_array_year_by_year_div_growth.append(input_champ_list[company])

    return champ_array_year_by_year_div_growth


def color_champ_list_after_year_by_year_div_growth_analysis(input_champ_list, hist_worksheet, color_fill):
    """"""

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


"""def binary_search(company_name, worksheet, first_indx, last_indx):
    low = first_indx
    high = last_indx

    while low <= high:
        mid = (low + high)//2
        #Company name
        guess = worksheet['A' + str(mid)].value
        print("Guess ", guess)
        print("company_name ", company_name)
        print("mid ", mid)
        print(guess < company_name)
        if guess == company_name:
            return mid
        if guess > company_name:
            high = mid - 1
        else:
            print("LOW")
            low = mid - 1
            print("LOW = ", low)
    return None
"""
