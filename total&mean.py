import pandas as pd

# read the file:
data = pd.read_excel('data1.xlsm',sheet = 'data')

Sender = data.Sender
Receiver = data.Receiver
Event = data.Event
Volume = data.Volume
Replies = data.Replies
Biz = data.Biz
Casual = data.Casual
Kindness = data.Kindness
Emotional = data.Emotional

# 1. For mean
peoemo = data.loc[Event == 'emotional'][['Sender','Event','Replies']]
meanpeoemo = peoemo.groupby(['Sender']).mean()

peokin = data.loc[Event == 'kindness'][['Sender','Event','Replies']]
meanpeokin = peokin.groupby(['Sender']).mean()

peobus = data.loc[Event == 'business'][['Sender','Event','Replies']]
meanpeobus = peobus.groupby(['Sender']).mean()


# 2. For Total:

# Find all rows that useful = 1
busSender = data.loc[Kindness == 1][['Sender', 'Receiver','Biz']]
# Find all rows that useful = -1
busReceiver = data.loc[Kindness == -1][['Sender', 'Receiver','Biz']]
# Find all rows that useful = 1
kinSender = data.loc[Kindness == 1][['Sender', 'Receiver','Kindness']]
# Find all rows that useful = -1
kinReceiver = data.loc[Kindness == -1][['Sender', 'Receiver','Kindness']]
# Find all rows that emotion = 1
emoSender = data.loc[Emotional == 1] [['Sender', 'Receiver','Emotional']]
# Find all rows that emotion = -1
emoReceiver = data.loc[Emotional == -1][['Sender', 'Receiver','Emotional']]
# Find all rows that casual = 1
casSender = data.loc[Casual == 1] [['Sender', 'Receiver','Casual']]
# Find all rows that casual = -1
casReceiver = data.loc[Casual == -1] [['Sender', 'Receiver','Casual']]

# (1). For use = 1 data 'Sender' refers to the Sender and data 'Receiver' refers to the Reciver
busSender1 = list(busSender.Sender)

# (2). For use = -1 data 'Sender' refers to the Receiver and data 'Receiver' refers to the Sender
busSender2 = list(busReceiver.Receiver)

# (3). for use(useful infomation), combine all senders together and combine all recivers together(use = 1 and use = -1):
allbusSender = busSender1 + busSender2
df = pd.DataFrame({'allbusSender':allbusSender})
countbus = df['allbusSender'].value_counts()

# (1). For use = 1 data 'Sender' refers to the Sender and data 'Receiver' refers to the Reciver
emoSender1 = list(emoSender.Sender)

# (2). For use = -1 data 'Sender' refers to the Receiver and data 'Receiver' refers to the Sender
emoSender2 = list(emoReceiver.Receiver)

# (3). for use(useful infomation), combine all senders together and combine all recivers together(use = 1 and use = -1):
allemoSender = emoSender1 + emoSender2
df = pd.DataFrame({'allemoSender':allemoSender})
countemo = df['allemoSender'].value_counts()


# (1). For use = 1 data 'Sender' refers to the Sender and data 'Receiver' refers to the Reciver
kinSender1 = list(kinSender.Sender)

# (2). For use = -1 data 'Sender' refers to the Receiver and data 'Receiver' refers to the Sender
kinSender2 = list(kinReceiver.Receiver)

# (3). for use(useful infomation), combine all senders together and combine all recivers together(use = 1 and use = -1):
allkinSender = kinSender1 + kinSender2
df = pd.DataFrame({'allkinSender':allkinSender})
countkin = df['allkinSender'].value_counts()


# (1). For use = 1 data 'Sender' refers to the Sender and data 'Receiver' refers to the Reciver
casSender1 = list(casSender.Sender)

# (2). For use = -1 data 'Sender' refers to the Receiver and data 'Receiver' refers to the Sender
casSender2 = list(casReceiver.Receiver)

# (3). for use(useful infomation), combine all senders together and combine all recivers together(use = 1 and use = -1):
allcasSender = casSender1 + casSender2
df = pd.DataFrame({'allcasSender':allcasSender})
countcas = df['allcasSender'].value_counts()


# Final ressult
result = pd.concat([meanpeobus,meanpeokin,meanpeoemo,countbus,countemo,countkin,countcas], axis = 1)
writer = pd.ExcelWriter('output.xlsx')
result.to_excel(writer,'Sheet1')
writer.save()
