from konlpy.tag import Kkma
import pandas as pd
from collections import Counter
from datetime import datetime, timedelta
from tqdm import tqdm


def date_range(start, end):
    start = datetime.strptime(start, "%Y%m%d")
    end = datetime.strptime(end, "%Y%m%d")
    dates = []
    for i in range((end - start).days + 1):
        dates.append((start + timedelta(i)).strftime("%Y%m%d"))
    return dates


start = input("START_DATE : ")
end = input("END_DATE : ")
dates = date_range(start, end)

m=0
count1 = 0
data1 = []
for i in dates: 
    print(f"Now_Date Roading : {i}")
    data = pd.read_excel(f"기사/{i}.xlsx")
    for j in data["기사 내용"]:
        data1.append(j)
        count1 += 1

print(count1)

noun1 = []
kkma = Kkma()
for i in tqdm(data1):
    noun = kkma.nouns(i)

    for j , v in enumerate(noun):
        if len(v)<2:
          noun.pop(j)
    
    for j in noun:
        noun1.append(j)

count = Counter(noun1)


noun_list = count.most_common(50)

df = pd.DataFrame(noun_list)
df.to_excel(f"결과/{start}~{end}.xlsx")
