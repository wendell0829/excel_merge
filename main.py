from excel_merge.core import load_to_dict, merge_workbook_dict, write_dict_to_excel


filename = ['files/test1.xlsx', 'files/test2.xlsx']

dict1 = load_to_dict(filename[0])
dict2 = load_to_dict(filename[1])
dict3 = merge_workbook_dict(dict1, dict2)
print(dict3)
write_dict_to_excel(dict3, filename='new1')