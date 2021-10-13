def first_company_in_list_cell(worksheet):
    """35 in the cycle was set experimentally just to make sure that we could find the first company eventually"""
    for i in range(1, 35):
        if worksheet['A' + str(i)].value == 'Name':
            return i + 1


def last_company_in_list_cell(worksheet):
    for i in range(500, 1, -1):
        if worksheet['A' + str(i)].value == 'Averages for All':
            return i - 2


def last_company_in_hist_list_cell(worksheet_hist):
    for i in range(1000, 1, -1):
        if worksheet_hist['B' + str(i)].value is not None:
            return i


#TODO sorting column in a worksheet, binary search
def find_company_in_list(ticker, worksheet, first_indx, last_indx):
    for i in range(first_indx, last_indx):
        if worksheet['B' + str(i)].value == ticker:
            return i
    return None


def to_fixed(numObj, digits=0):
    """Truncate real numbers from 2.98456 to 2.98"""
    return f"{numObj:.{digits}f}"

