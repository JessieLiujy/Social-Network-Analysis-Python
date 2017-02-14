import pandas as pd

# read the file:
data = pd.read_excel('data_master.xlsm',sheet = 'ws1')

Sender = data.联系人
Message = data.消息
Time = data.时间
data.head(5)

Senderlist = list(Sender)
indexes=[Senderlist.index(l) for l in Senderlist] # assign index as ID number

data.insert(2,"ID",indexes) # add the ID column to the dataset
data.head() # check the data
data[data.ID==1] # Example: Select all the data with Sender ID = 1

writer = pd.ExcelWriter('data with ID.xlsx')
data.to_excel(writer, index=False)
writer.save()
