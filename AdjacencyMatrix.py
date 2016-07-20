# Python 3.5.1

from __future__ import print_function
import itertools
import numpy as np
import pandas as pd

# read the file:
data = pd.read_excel('data-example.xlsx')

# check the data: data.head(5)

# define the name of the data
Sender = data.Sender
Receiver = data.Receiver
Useful = data.Useful
Emotion = data.Emotion
Casual = data.Casual

# 1. Find all data:
# Find all rows that useful = 1
useSender = data.loc[Useful == 1][['Sender', 'Receiver','Useful']]
# Find all rows that useful = -1
useReceiver = data.loc[Useful == -1][['Sender', 'Receiver','Useful']]
# Find all rows that emotion = 1
emoSender = data.loc[Emotion == 1] [['Sender', 'Receiver','Emotion']]
# Find all rows that emotion = -1
emoReceiver = data.loc[Emotion == -1][['Sender', 'Receiver','Emotion']]
# Find all rows that casual = 1
casSender = data.loc[Casual == 1] [['Sender', 'Receiver','Casual']]
# Find all rows that casual = -1
casReceiver = data.loc[Casual == -1] [['Sender', 'Receiver','Casual']]

# 2. Define Sender and Receiver to make the matrix:
# 2.1 for useful:
# (1). For use = 1 data, Sender refers to the provider of useful information
useSender1 = list(useSender.Sender)
useReceiver1 = list(useSender.Receiver)

# (2). For use = -1 data, Receiver refers to the provider of useful information
useSender2 = list(useReceiver.Receiver)
useReceiver2 = list(useReceiver.Sender)

# (3). for all use = 1 and use = -1, combine all senders(providers) together and combine all recivers together:
alluseSender = useSender1 + useSender2
alluseReceiver = useReceiver1 + useReceiver2
# (4). Put the data in to format (Sender, Receiver) -- one to one correspondence
useedges = list(zip(alluseSender,alluseReceiver))

# 2.2 For emotional(same way):
# For emotion = 1
emoSender1 = list(emoSender.Sender)
emoReceiver1 = list(emoSender.Receiver)
# For emotion = -1
emoSender2 = list(emoReceiver.Receiver)
emoReceiver2 = list(emoReceiver.Sender)

# combine data that emotion = 1 and emotion = -1 together:
allemoSender = list(emoSender1) + list(emoSender2)
allemoReceiver = list(emoReceiver1) + list(emoReceiver2)
emoedges = list(zip(allemoSender,allemoReceiver))

# 2.3 For casual(same way):
# For casual = 1
casSender1 = list(casSender.Sender)
casReceiver1 = list(casSender.Receiver)
# For casual = -1
casSender2 = list(casReceiver.Receiver)
casReceiver2 = list(casReceiver.Sender)

# combine data that casual = 1 and casual = -1 together:
allcasSender = casSender1 + casSender2
allcasReceiver = casReceiver1 + casReceiver2
casedges = list(zip(allcasSender,allcasReceiver))

# 3. arrays for the matrixs:
# Get the name list for all people to create the column and row for the matrix
n = useedges + emoedges + casedges
allpeople = sorted(list(set((itertools.chain.from_iterable(n)))))

# 3.1 For useful:
# edges:
usecount = dict((x,useedges.count(x)) for x in set(useedges))

uselist = list()
for row in allpeople:
	for column in allpeople:
		uselist.append(usecount.get((row, column), 0))
datause = [uselist[i:i+len(allpeople)] for i in range(0, len(uselist), len(allpeople))]

# 3.2 For emotional:
emocount = dict((x,emoedges.count(x)) for x in set(emoedges))
emolist = list()
for row in allpeople:
	for column in allpeople:
		emolist.append(emocount.get((row, column), 0))
dataemo = [emolist[i:i+len(allpeople)] for i in range(0, len(emolist), len(allpeople))]

# 3.3 for casual:
cascount = dict((x,casedges.count(x)) for x in set(casedges))
allcas = set(allcasSender+allcasReceiver)
caslist = list()
for row in allpeople:
	for column in allpeople:
		caslist.append(cascount.get((row, column), 0))
datacas = [caslist[i:i+len(allpeople)] for i in range(0, len(caslist), len(allpeople))]

# 4. export to the excel file:
dfuse = pd.DataFrame(datause, index = allpeople, columns = allpeople) # senders in colunm and receiver in row

writer = pd.ExcelWriter('output_matrix.xlsx')
dfuse.to_excel(writer, sheet_name = 'use')

dfemo = pd.DataFrame(dataemo, index = allpeople, columns = allpeople)
dfemo.to_excel(writer, sheet_name = 'emotional')

dfcas = pd.DataFrame(datacas, index = allpeople, columns = allpeople)
dfcas.to_excel(writer, sheet_name = 'casual')

writer.save()






            


    



    
            

    

