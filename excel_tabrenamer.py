### Variables
input = "/Users/Olivia/Desktop/fools.xlsx"

### Code
def tab_renamer (input):

    ### Imports
    import openpyxl as xl
    from pathlib import Path

    ### Open Excel Workbook
    wb = xl.load_workbook(input)


    ### Create List of New Names based on new_names tab
    new_names = []
    sheet = wb['new_names']
    for row in sheet.rows:
        new_names.append(row[0].value)

    print(new_names)

    ## Change Tab Names
    old_names = wb.sheetnames
    old_names.remove("new_names") # Get List of Current Tab Names

    print(old_names)

    if len(old_names) == len(new_names):
        i = 0
        for old_name in old_names:
            wb[old_name].title = new_names[i]
            i+=1

        wb.save(str(Path(input).parent/Path(input).stem) + '_edited' + str(Path(input).suffix))

    else:
        print('Number of New Names don\'t match Number of Tabs')

### Run Code
tab_renamer (input)
