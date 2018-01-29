

import pandas as pd

x1 = pd.ExcelFile('Boundary_Train.xlsx')

df = x1.parse("Sheet1")

df_=pd.read_excel('Boundary_Train.xlsx',sheetname='Sheet1')
# example of valid data
# df.values[] starts validated data from index 2
# df.values[2] as ['instruments', 'instruments number',...]
# the true data starts from index 3
# df.values[3] as array([u'BAB', u'Instrument 1', 1, 0, 0, 0, 0, 0, 0, 0, 0, u'trading book'], dtype=object)

data = df.values
start_col = 0
start_row = 0
search_flag = True
result = [''] * (len(df.index)) 

# find the valid data beginning and end of row and column
# each row
for i in range(len(df.index)):
    # each column
    for j in range(len(data[i])):
        if data[i][j] == 'Trading':
            start_col = j
            start_row = i + 1
            search_flag = False # stop searching
            break
    if search_flag == False:
        break
#print 'start row and start col:', start_row, start_col

# start making the decision
for i in range(start_row,len(df.index)):
    if data[i][start_col] == 1 and sum(data[i][start_col + 1:]) == 0:
        result[i] = 'trading book'
#        print(result[i])
    elif data[i][start_col]+ data[i][start_col + 1] == 2 or data[i][start_col] + data[i][start_col + 1] == 0:
        result[i] = 'Error'
#        print(result[i])
    else:
        result[i] = 'banking book'
#        print(result[i])

##################################
# export to the original file


df = pd.DataFrame(data)
new_result = pd.DataFrame([data[:,1],result]).transpose()
writer = pd.ExcelWriter('output.xlsx')
new_result.to_excel(writer,'Sheet1')
writer.save()

###################################
## for test and check part#############
#for i in range(start_row,df.index.max() + 1):
#    if data[i][-1] != result[i]:
#        print('error on row,' + chr(i))
##        print 'error on row:', i
##        print  'true result:', data[i][-1], 'computed result:', result[i]
##        print sum(data[i][start_col + 1:-1])
##        break
##print result[start_row:50]
###################################
