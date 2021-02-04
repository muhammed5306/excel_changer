# import pandas as pd
# from numpy.core.tests.test_umath_accuracy import convert
#
# asd = pd.read_excel("C:\\Users\\MSI\\Desktop\\Excel_Chan\\excel\\Company.xlsx", dtype='str')
# # asd = pd.read_excel("C:\\Users\\MSI\\Desktop\\excel\\Company.xlsx", index_col=1, header=0, usecols=['ID','Adı'])
#
# # print(asd)
#
# column = asd.columns.ravel()
#
#
# id_col = asd[column[0]].tolist()
# # id_col_int = asd[column[0]].astype('int')
# selected_col = asd[column[1]].tolist()
#
# zlist = []
#
# for k in id_col:
#     if k == id_col:
#         print(k)
#
# new_list = []

# for i in id_col:
#     print(i)
#     if type(i) == int or float:
#         zlist.append(i)
# #
# for i, j in zip(id_col,selected_col):
#     data = f'({i},{j})'
#     new_list.append(data)


# for i, j in zip(id_col,selected_col):
#     new_list.append([i,j])
#
#
# print("ID Column",id_col)
# print("Selected Column", selected_col)
# # for

# list1 = [['1', 'apple'], ['2', 'b'], ['3', 'cell']]
list2 = ['Apple Tech.', 'D', 'nan', 'nan', 'E', 'nan']
list3 = [1, 2, 3]


# string = list2
# a = str(string).replace(f"{list2[0]}",f"{list3[0]}")
# print(a)
kd = {}
j = 0
for i in range(0,len(list2)):
    if list2[i] != 'nan':

        kd.update({list2[i]: list3[j]})
        if j < len(list3):
            j = j + 1

        else:
            break



print(kd)
list2 = [kd.get(x, x) for x in list2]
print(list2)


# print(list2)
# for n, i in enumerate(a):

# for i in range(0,len(list1)):
#     for j in range(0,len(list2)):
#         if str(list2[j]).lower().find(list1[i][1]) != -1:
#             print(list1[i][1])
#             print(list2[j])

        # if str(list2[j]).find(list1[i][1]) != -1:
        #     print(list1[i][1])
        #     print(list2[j])

# print(new_list[0][1])

# for i in range(0,len(new_list)):
#     print(new_list[i])



# print(zlist)
# print(selected_col)
# print(new_list)

# listasd = []

# print(id_col,selected_col)
