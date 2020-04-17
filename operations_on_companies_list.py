def first_company_in_list_cell(worksheet):
    for i in range(1, 35):
        if worksheet['A' + str(i)].value == 'Name':
            return i + 1


def last_company_in_list_cell(worksheet):
    for i in range(500, 1, -1):
        if worksheet['A' + str(i)].value == 'Averages for All':
            return i - 2


def last_company_in_hist_list_cell(worksheet_hist):
    for i in range(1000, 1, -1):
        if worksheet_hist['A' + str(i)].value is not None:
            return i


#TODO sorting column in a worksheet, binary search
def find_company_in_list(company_name, worksheet, first_indx, last_indx):
    for i in range(first_indx, last_indx):
        if worksheet['A' + str(i)].value == company_name:
            return i
    return None


def to_fixed(numObj, digits=0):
    """Truncate real numbers from 2.98456 to 2.98"""
    return f"{numObj:.{digits}f}"


def get_fundamental_analysis_on_champs_list(ws, eval_model):
    """Analysing preliminary list of potential champs in stock list using fundamental analysis parameters"""

    print("Analysing preliminary champions list")
    prelim_champ_list = []
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
            prelim_champ_list.append(champion)

    return prelim_champ_list


#def get_colored_fundamental_parameters_of_company_in_list():


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