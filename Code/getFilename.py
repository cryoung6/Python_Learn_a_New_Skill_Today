''' getFilename.py

    Create the path to the Results subfolder on my MAC as "mypath"

    Get the current date and time and convert to a string.
    For example, "d1" will be 04302021   08_19_AM

    Return a string of :
        the path and new filename (d1 + 'Business Expenses')

'''


def getFilename():
    from datetime import datetime
    import os

    mypath = os.environ['HOME']
    mypath += '/BOOK_AUTHORING/Python_Lab2/CODE_for_Lab2/results/'
    d1 = datetime.strftime(datetime.now(), '%m%d%Y   %H_%M_%p')
    return mypath + d1 + ' Business Expenses.xlsx'
