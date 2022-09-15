import os
import xlrd
from xlutils.copy import copy
def addp():
    dir = r"./test1/"
    osfile = []
    for root, dirs, files in os.walk(dir):
        for file in files:
            if file.split('.')[1] == 'txt':
                osfile.append(file)
    for indexFile in range(len(osfile)):
        with open('./test1/' + osfile[indexFile], "r", encoding="utf-8") as f:
            lines = f.readlines()
    count = 0
    for index in range(len(lines)):
        if lines[index] == '.\n':
            count+=1
    print(count)
    print(lines)

def extract():
    dir = r"./test1/"
    osfile = []
    for root, dirs, files in os.walk(dir):
        for file in files:
            if file.split('.')[1] == 'txt':
                osfile.append(file)
    for indexFile in range(len(osfile)):
        with open('./test1/' + osfile[indexFile], "r", encoding="utf-8") as f:
            lines = f.readlines()
            # print(lines)
            flatstrat = []
            flatend = []
            for index in range(len(lines)):
                if lines[index].find('skos:prefLabel')!=-1:
                    flatstrat.append(index-1)
            # 由于每一段结尾并不能确定，所有采用flatstrat.append(index-5)当每一段结尾，最后一个结尾为len(lines)
            for end in range(1,len(flatstrat)):
                flatend.append(flatstrat[end]-2)
            flatend.append(len(lines))
            #print(len(flatstrat),len(flatend))
            # 将没段存入词典
            examples = []
            for indexSE in range(len(flatstrat)):
                examples.append(lines[flatstrat[indexSE]:flatend[indexSE]+1])

            # 清洗
            count = 0
            for exs in range(len(examples)):
                results.append([x.strip() for x in examples[exs] if x.strip() != ''])
            for res in range(len(results)):
                results[res].pop()
                for resinner in range(len(results[res])):
                    results[res][resinner]=results[res][resinner].replace(';','').replace('@en','').replace('early','Lower').replace('late','Upper').replace('middle','Middle').strip()
    # 锚点检测一下一定有skos:prefLabel 但不一定有utf:partOf因此要达到数量一一致
    for indexre in range(len(results)):
        flagpartOf = 1
        for indexreinner in range(len(results[indexre])):
            if results[indexre][indexreinner].find('utf:partOf') != -1:
                flagpartOf = 0
        if flagpartOf:
            results[indexre].append('utf:partOf utf:NUlL')

    # 提取所有skos:prefLabel 和 utf:partOf 的数据
    for indexre in range(len(results)):
        for indexreinner in range(len(results[indexre])):
            if results[indexre][indexreinner].find('prefLabel') != -1:
                prefLabel.append(results[indexre][indexreinner])
            if results[indexre][indexreinner].find('utf:partOf') != -1:
                partOf.append(results[indexre][indexreinner])
    for index in range(len(results)):
        print(results[index])
    print('===========================================================')
if __name__ == '__main__':
    addp()
    results = []
    prefLabel = []
    partOf = []
    #extract()
    print(len(results))
    print(len(prefLabel))
    print(len(partOf))