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