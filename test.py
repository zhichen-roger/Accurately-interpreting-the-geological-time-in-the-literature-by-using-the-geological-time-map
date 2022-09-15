import os
import xlrd
from xlutils.copy import copy
def extract():
    dir = r"./test/"
    osfile = []
    for root, dirs, files in os.walk(dir):
        for file in files:
            if file.split('.')[1] == 'txt':
                osfile.append(file)
    for indexFile in range(len(osfile)):
        with open('./test/' + osfile[indexFile], "r", encoding="utf-8") as f:
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
    # for index in range(len(results)):
    #     print(results[index])
    print('===========================================================')
def targetxlsx():
    # 打开文件，获取excel文件的workbook（工作簿）对象
    workbook = xlrd.open_workbook("test.xlsx")  # 文件路径
    worksheet = workbook.sheet_by_index(0)
    name = worksheet.name  # 获取表的姓名
    ncols = worksheet.ncols  # 获取该表总列数
    # 找到所有列数据
    for cols in range(ncols):
        # print(worksheet.col_values(cols))
        # print(len(worksheet.col_values(index)))
        # 用数组进行备份
        totalTagert.append(worksheet.col_values(cols))
    # for totalIndex in range(len(totalTagert)):
    #     print(totalTagert[totalIndex])

def targetMatch():
    # 打印目标三列
    LMU = totalTagert[10]
    print(LMU)
    Period = totalTagert[11]
    print(Period)
    Stage = totalTagert[12]
    print(Stage)
    print('====+++++++++++++++++++++++++++=====')
    print(prefLabel)
    print(partOf)
    for indexstage in range(1,len(Stage)):
        # 第一种情况Stage存在的时候且Period和LMU都有
        if Stage[indexstage] and Period[indexstage] and LMU[indexstage]:
            flagindex = 1
            innerindex = []
            for indexpref in range(len(prefLabel)):
                if prefLabel[indexpref].lower().find(Stage[indexstage].lower())!= -1 and partOf[indexpref].lower().find(Period[indexstage].lower())!= -1:
                    # 找到了
                    flagindex = 0
                    innerindex.append(indexpref)
                    # print(indexpref)
                    # print(Stage[indexstage])
                    # print(Period[indexstage])
                    # print(prefLabel[indexpref])
                    # print(partOf[indexpref])
                    # print('=====================================')
            # 没找到
            if flagindex:
                innerindex.append(-1)
            indexresults.append(innerindex)
        # 第二种情况Period和LMU都有且LMU不为缩写
        elif Period[indexstage] and not LMU[indexstage].find('-')!= -1:
            flagindex = 1
            innerindex = []
            for indexpref in range(len(prefLabel)):
                if prefLabel[indexpref].lower().find(Period[indexstage].lower()) != -1 and prefLabel[indexpref].lower().find(LMU[indexstage].lower()) != -1:
                    # 找到了
                    flagindex = 0
                    innerindex.append(indexpref)
                # elif prefLabel[indexpref].lower().find(Period[indexstage].lower()) != -1:
                #     # 找到了
                #     flagindex = 0
                #     innerindex.append(indexpref)
            # 没找到
            if flagindex:
                innerindex.append(-1)
            indexresults.append(innerindex)
        # 第三种情况Period和LMU且LMU为缩写模式
        elif Period[indexstage] and LMU[indexstage].find('-')!=-1:
            flagindex = 1
            innerindex = []
            # ('early','Lower') ('late','Upper') ('middle','Middle') L - M - U
            if LMU[indexstage].find('L') != -1 and LMU[indexstage].find('U') != -1:
                # print(LMU[indexstage])
                innerindex.append(-2)
                for indexpref in range(len(prefLabel)):
                    if prefLabel[indexpref].lower().find(Period[indexstage].lower()) != -1 and (prefLabel[
                        indexpref].lower().find('Lower'.lower()) != -1 or prefLabel[
                        indexpref].lower().find('Upper'.lower()) != -1):
                        # 找到了
                        flagindex = 0
                        innerindex.append(indexpref)
            elif LMU[indexstage].find('L') != -1 and LMU[indexstage].find('M')!= -1 :
                #print(LMU[indexstage])
                innerindex.append(-2)
                for indexpref in range(len(prefLabel)):
                    if prefLabel[indexpref].lower().find(Period[indexstage].lower()) != -1 and (prefLabel[
                        indexpref].lower().find('Lower'.lower()) != -1 or prefLabel[
                        indexpref].lower().find('Middle'.lower()) != -1):
                        # 找到了
                        flagindex = 0
                        innerindex.append(indexpref)
            elif LMU[indexstage].find('M')!= -1 and LMU[indexstage].find('U') != -1:
                print(LMU[indexstage])
                innerindex.append(-2)
                for indexpref in range(len(prefLabel)):
                    if prefLabel[indexpref].lower().find(Period[indexstage].lower()) != -1 and (prefLabel[
                        indexpref].lower().find('Middle'.lower()) != -1 or prefLabel[
                        indexpref].lower().find('Upper'.lower()) != -1):
                        # 找到了
                        flagindex = 0
                        innerindex.append(indexpref)
            # 没找到
            if flagindex:
                innerindex.append(-1)
            indexresults.append(innerindex)
        else:
            indexresults.append([-1])

def writeToTarger(path, writeresult):
    #print(indexresults)
    for index in range(len(indexresults)):
        print(indexresults[index])
    # 读取存在数字列表如果是唯一一个数据，且不是空数据
    for index in range(len(indexresults)):
        wirtelist = []
        if len(indexresults[index]) == 1 and indexresults[index][0] != -1:
            #print(results[indexresults[index][0]])
            for writeindex in range(len(results[indexresults[index][0]])):
                if results[indexresults[index][0]][writeindex].find('utf:start ') != -1:
                    wirtelist.append(results[indexresults[index][0]][writeindex][9:].strip())
                if results[indexresults[index][0]][writeindex].find('utf:start_uncertainty ') != -1:
                    wirtelist.append(results[indexresults[index][0]][writeindex][21:].strip())
                if results[indexresults[index][0]][writeindex].find('utf:end ') != -1:
                    wirtelist.append(results[indexresults[index][0]][writeindex][7:].strip())
                if results[indexresults[index][0]][writeindex].find('utf:end_uncertainty ') != -1:
                    wirtelist.append(results[indexresults[index][0]][writeindex][19:].strip())
                if results[indexresults[index][0]][writeindex].find('utf:reference_system utf:') != -1:
                    wirtelist.append(results[indexresults[index][0]][writeindex][25:].strip())
            writeresult.append(wirtelist)
        elif len(indexresults[index]) != 1 and indexresults[index][0] != -2 and indexresults[index][0] != -1:
            # print(indexresults[index])
            # 建立temp比大小 默认为第一个
            starttemp = ''
            start_uncertaintytemp = ''
            flotstart_uncertaintytemp = 0.0
            endtemp = ''
            end_uncertaintytemp = ''
            referencetemp = ''
            for writeindex in range(len(results[indexresults[index][0]])):
                if results[indexresults[index][0]][writeindex].find('utf:start ') != -1:
                    starttemp = results[indexresults[index][0]][writeindex][9:].strip()
                if results[indexresults[index][0]][writeindex].find('utf:start_uncertainty ') != -1:
                    start_uncertaintytemp = results[indexresults[index][0]][writeindex][21:].strip()
                    flotstart_uncertaintytemp = float(start_uncertaintytemp)
                if results[indexresults[index][0]][writeindex].find('utf:end ') != -1:
                    endtemp  = results[indexresults[index][0]][writeindex][7:].strip()
                if results[indexresults[index][0]][writeindex].find('utf:end_uncertainty ') != -1:
                    end_uncertaintytemp = results[indexresults[index][0]][writeindex][19:].strip()
                if results[indexresults[index][0]][writeindex].find('utf:reference_system utf:') != -1:
                    referencetemp = results[indexresults[index][0]][writeindex][25:].strip()
            # 从第二个开始
            for indexinner in range(1,len(indexresults[index])):
                # print(results[indexresults[index][indexinner]])
                start = ''
                start_uncertainty = ''
                floatstart_uncertainty = 0.0
                end = ''
                end_uncertainty = ''
                reference = ''
                for indextemp in range(len(results[indexresults[index][indexinner]])):
                    if results[indexresults[index][indexinner]][indextemp].find('utf:start_uncertainty ') != -1:
                        start_uncertainty = results[indexresults[index][indexinner]][indextemp][21:].strip()
                        floatstart_uncertainty = float(start_uncertainty)
                    if results[indexresults[index][indexinner]][indextemp].find('utf:start ') != -1:
                        start = results[indexresults[index][indexinner]][indextemp][9:].strip()
                    if results[indexresults[index][indexinner]][indextemp].find('utf:end ') != -1:
                        end = results[indexresults[index][indexinner]][indextemp][7:].strip()
                    if results[indexresults[index][indexinner]][indextemp].find('utf:end_uncertainty ') != -1:
                        end_uncertainty = results[indexresults[index][indexinner]][indextemp][19:].strip()
                    if results[indexresults[index][indexinner]][indextemp].find('utf:reference_system utf:') != -1:
                        reference = results[indexresults[index][indexinner]][indextemp][25:].strip()
                # 比大小
                if flotstart_uncertaintytemp > floatstart_uncertainty:
                    starttemp = start
                    start_uncertaintytemp = start_uncertainty
                    endtemp = end
                    end_uncertaintytemp = end_uncertainty
                    referencetemp = reference
            wirtelist.append(starttemp)
            wirtelist.append(start_uncertaintytemp)
            wirtelist.append(endtemp)
            wirtelist.append(end_uncertaintytemp)
            wirtelist.append(referencetemp)
            writeresult.append(wirtelist)
        #######################################################################
        # 第三种输入中存在多个LMU的
        elif len(indexresults[index]) != 1 and indexresults[index][0] == -2:
            indexresults[index].pop(0)
            print(indexresults[index])
            L = []
            M = []
            U = []
            # 将LMU各项分类
            for indexLMU in range(len(indexresults[index])):
                if results[indexresults[index][indexLMU]][1].lower().find('Lower'.lower())!=-1:
                    L.append(results[indexresults[index][indexLMU]])
                if results[indexresults[index][indexLMU]][1].lower().find('Upper'.lower())!=-1:
                    U.append(results[indexresults[index][indexLMU]])
                if results[indexresults[index][indexLMU]][1].lower().find('Middle'.lower())!=-1:
                    M.append(results[indexresults[index][indexLMU]])
            # 建立分别比较原则 L-U start start_uncertainty为L,end end_uncertainty为U,因此依次分为3类
            print('&'*100)
            print(L)
            print(U)
            print(M)
            print('+'*100)
            # 建立通用保存数据数列
            starttemp = ''
            start_uncertaintytemp = ''
            flotstart_uncertaintytemp = 0.0
            endtemp = ''
            end_uncertaintytemp = ''
            floatend_uncertaintytemp = 0.0
            referencetemp = ''
            # 第一类L-U
            if L and U:
                # 比较晒选出符合L的start start_uncertainty
                # 建立temp比大小 默认为第一个
                for indexL in range(len(L[0])):
                    if L[0][indexL].find('utf:start ') != -1:
                        starttemp = L[0][indexL][9:].strip()
                    if L[0][indexL].find('utf:start_uncertainty ') != -1:
                        start_uncertaintytemp = L[0][indexL][21:].strip()
                        flotstart_uncertaintytemp = float(start_uncertaintytemp)
                    if L[0][indexL].find('utf:reference_system utf:') != -1:
                        referencetemp = L[0][indexL][25:].strip()
                for indexU in range(len(U[0])):
                    if U[0][indexU].find('utf:end ') != -1:
                        endtemp = U[0][indexU][7:].strip()
                    if U[0][indexU].find('utf:end_uncertainty ') != -1:
                        end_uncertaintytemp = U[0][indexU][19:].strip()
                        floatend_uncertaintytemp = float(end_uncertaintytemp)
                #从第二个开始 L 为start U 为end
                for indexLt in range(1,len(L)):
                    start = ''
                    start_uncertainty = ''
                    floatstart_uncertainty = 0.0
                    reference = ''
                    for indexLttemp in range(len(L[indexLt])):
                        if L[indexLt][indexLttemp].find('utf:start ') != -1:
                            start = L[indexLt][indexLttemp][9:].strip()
                        if L[indexLt][indexLttemp].find('utf:start_uncertainty ') != -1:
                            start_uncertainty =  L[indexLt][indexLttemp][21:].strip()
                            floatstart_uncertainty = float(start_uncertainty)
                        if L[indexLt][indexLttemp].find('utf:reference_system utf:') != -1:
                            reference = L[indexLt][indexLttemp][25:].strip()
                    # 比大小
                    if flotstart_uncertaintytemp > floatstart_uncertainty:
                        starttemp = start
                        start_uncertaintytemp = start_uncertainty
                        referencetemp = reference
                for indexUt in range(1,len(U)):
                    end = ''
                    end_uncertainty = ''
                    floatend_uncertainty = 0.0
                    for indexUttemp in range(len(U[indexUt])):
                        if U[indexUt][indexUttemp].find('utf:end ') != -1:
                            end = U[indexUt][indexUttemp][7:].strip()
                        if U[indexUt][indexUttemp].find('utf:end_uncertainty ') != -1:
                            end_uncertainty =  U[indexUt][indexUttemp][19:].strip()
                            floatend_uncertainty = float(end_uncertainty)
                    # 比大小
                    if floatend_uncertaintytemp > floatend_uncertainty:
                        endtemp = end
                        end_uncertaintytemp = end_uncertainty
            # 第二类L-M
            elif L and M:
                # 比较晒选出符合L的start start_uncertainty
                # 建立temp比大小 默认为第一个
                for indexL in range(len(L[0])):
                    if L[0][indexL].find('utf:start ') != -1:
                        starttemp = L[0][indexL][9:].strip()
                    if L[0][indexL].find('utf:start_uncertainty ') != -1:
                        start_uncertaintytemp = L[0][indexL][21:].strip()
                        flotstart_uncertaintytemp = float(start_uncertaintytemp)
                    if L[0][indexL].find('utf:reference_system utf:') != -1:
                        referencetemp = L[0][indexL][25:].strip()
                for indexM in range(len(M[0])):
                    if M[0][indexM].find('utf:end ') != -1:
                        endtemp = M[0][indexM][7:].strip()
                    if M[0][indexM].find('utf:end_uncertainty ') != -1:
                        end_uncertaintytemp = M[0][indexM][19:].strip()
                        floatend_uncertaintytemp = float(end_uncertaintytemp)
                # 从第二个开始 L 为start M 为end
                for indexLt in range(1, len(L)):
                    start = ''
                    start_uncertainty = ''
                    floatstart_uncertainty = 0.0
                    reference = ''
                    for indexLttemp in range(len(L[indexLt])):
                        if L[indexLt][indexLttemp].find('utf:start ') != -1:
                            start = L[indexLt][indexLttemp][9:].strip()
                        if L[indexLt][indexLttemp].find('utf:start_uncertainty ') != -1:
                            start_uncertainty = L[indexLt][indexLttemp][21:].strip()
                            floatstart_uncertainty = float(start_uncertainty)
                        if L[indexLt][indexLttemp].find('utf:reference_system utf:') != -1:
                            reference = L[indexLt][indexLttemp][25:].strip()
                    # 比大小
                    if flotstart_uncertaintytemp > floatstart_uncertainty:
                        starttemp = start
                        start_uncertaintytemp = start_uncertainty
                        referencetemp = reference
                for indexMt in range(1,len(M)):
                    end = ''
                    end_uncertainty = ''
                    floatend_uncertainty = 0.0
                    for indexMttemp in range(len(M[indexMt])):
                        if M[indexMt][indexMttemp].find('utf:end ') != -1:
                            end = M[indexMt][indexMttemp][7:].strip()
                        if M[indexMt][indexMttemp].find('utf:end_uncertainty ') != -1:
                            end_uncertainty =  M[indexMt][indexMttemp][19:].strip()
                            floatend_uncertainty = float(end_uncertainty)
                    # 比大小
                    if floatend_uncertaintytemp > floatend_uncertainty:
                        endtemp = end
                        end_uncertaintytemp = end_uncertainty
            # 第三类
            elif M and U:
                # 比较晒选出符合M的start start_uncertainty
                # 建立temp比大小 默认为第一个
                for indexM in range(len(M[0])):
                    if M[0][indexM].find('utf:start ') != -1:
                        starttemp = M[0][indexM][9:].strip()
                    if M[0][indexM].find('utf:start_uncertainty ') != -1:
                        start_uncertaintytemp = M[0][indexM][21:].strip()
                        flotstart_uncertaintytemp = float(start_uncertaintytemp)
                    if M[0][indexM].find('utf:reference_system utf:') != -1:
                        referencetemp = M[0][indexM][25:].strip()
                for indexU in range(len(U[0])):
                    if U[0][indexU].find('utf:end ') != -1:
                        endtemp = U[0][indexU][7:].strip()
                    if U[0][indexU].find('utf:end_uncertainty ') != -1:
                        end_uncertaintytemp = U[0][indexU][19:].strip()
                        floatend_uncertaintytemp = float(end_uncertaintytemp)
                # 从第二个开始 L 为start U 为end
                for indexMt in range(1,len(M)):
                    start = ''
                    start_uncertainty = ''
                    floatstart_uncertainty = 0.0
                    reference = ''
                    for indexMttemp in range(len(M[indexMt])):
                        if M[indexMt][indexMttemp].find('utf:start ') != -1:
                            start = M[indexMt][indexMttemp][9:].strip()
                        if M[indexMt][indexMttemp].find('utf:start_uncertainty ') != -1:
                            start_uncertainty =  M[indexMt][indexMttemp][21:].strip()
                            floatstart_uncertainty = float(start_uncertainty)
                        if M[indexMt][indexMttemp].find('utf:reference_system utf:') != -1:
                            reference = M[indexMt][indexMttemp][25:].strip()
                    # 比大小
                    if flotstart_uncertaintytemp > floatstart_uncertainty:
                        starttemp = start
                        start_uncertaintytemp = start_uncertainty
                        referencetemp = reference
                for indexUt in range(1,len(U)):
                    end = ''
                    end_uncertainty = ''
                    floatend_uncertainty = 0.0
                    for indexUttemp in range(len(U[indexUt])):
                        if U[indexUt][indexUttemp].find('utf:end ') != -1:
                            end = U[indexUt][indexUttemp][7:].strip()
                        if U[indexUt][indexUttemp].find('utf:end_uncertainty ') != -1:
                            end_uncertainty =  U[indexUt][indexUttemp][19:].strip()
                            floatend_uncertainty = float(end_uncertainty)
                    # 比大小
                    if floatend_uncertaintytemp > floatend_uncertainty:
                        endtemp = end
                        end_uncertaintytemp = end_uncertainty

            print(starttemp)
            print(start_uncertaintytemp)
            print(endtemp)
            print(end_uncertaintytemp)
            print(referencetemp)
            wirtelist.append(starttemp)
            wirtelist.append(start_uncertaintytemp)
            wirtelist.append(endtemp)
            wirtelist.append(end_uncertaintytemp)
            wirtelist.append(referencetemp)
            writeresult.append(wirtelist)
        else:
            writeresult.append(['NULL'])
    print('*'*100)
    for index in range(len(writeresult)):
        print(writeresult[index])
    print(len(writeresult))
    print('=============================')
    # 写入表格
    index = len(writeresult)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(writeresult[i])):
            new_worksheet.write(i + 1, j + 19, writeresult[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")
if __name__ == '__main__':
    results = []
    prefLabel = []
    partOf = []
    extract()

    totalTagert = []
    targetxlsx()

    LMU = []
    Period = []
    Stage = []
    TargetMatchIndex = []
    indexresults = []
    targetMatch()
    writeresult = []
    book_name_xls = 'test.xlsx'
    writeToTarger(book_name_xls, writeresult)