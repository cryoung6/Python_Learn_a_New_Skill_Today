''' getCategory.py

    Loop through the expCategories_Dict dictionary
    expCategories_Dict[phrase] = category

    The dictionary key is "phrase"

    If the "phrase" is in thisFileName
    return the category in the variable "thisCategory"

    Return "thisCategory"
'''


def getCategory(thisFileName, expCategories_Dict):
    thisCategory = ''
    for phrase in expCategories_Dict:
        if phrase in thisFileName.upper():
            thisCategory = expCategories_Dict[phrase]
            return thisCategory
    return thisCategory
