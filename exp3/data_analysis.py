from cnsenti import Sentiment
import csv

senti = Sentiment()
sentiment_list_pos = []
sentiment_list_neg = []
content = []
time = []
final_result = []
with open(r'FILEPATH','r',encoding='utf-8') as f:
    reader = csv.reader(f)
    for i in reader:
        # print(i)
        # print(i[4])
        # print(i[4].strip('TOPIC'))
        content.append(i[4].strip('TOPIC'))
        time.append(i[1][-6:-3])
        sentiment_list_pos.append(senti.sentiment_calculate(i[4].strip('TOPIC'))['pos'])
        sentiment_list_neg.append(senti.sentiment_calculate(i[4].strip('TOPIC'))['neg'])
print(sentiment_list_pos)
print(sentiment_list_neg)

for i in range(len(content)):
    temp_list = []
    temp_list.append(content[i])
    temp_list.append(sentiment_list_pos[i])
    temp_list.append(sentiment_list_neg[i])
    temp_list.append(time[i])
    final_result.append(temp_list)

with open(r'FILEPATH2','a+',encoding='utf-8',newline='') as file:
    writera = csv.writer(file)
    writera.writerows(final_result)
