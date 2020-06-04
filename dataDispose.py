
class Dispose:
    def __init__(self):
        pass

    def analyze(self,d):
        #newData
        newData = []
        for list in d:
            for l in list :
                if str(l[1]) ==  '1':
                    l[1] = '交货总成'
                    if len(l) > 34 and l[7] !='':
                        newData.append([l[1],l[7],l[10],l[11],l[31],str(l[17])+'+'+str(l[29]),'',str(l[33])+str(l[34])+'#','',l[12],l[27],l[28]])
                elif str(l[2]) == '2' and l[7] !='':
                    l[1] = '.1'
                    if len(l) > 34:
                        newData.append([l[1],l[7],l[10],l[11],l[31],str(l[17])+str(l[29]),'',str(l[33])+str(l[34]),'',l[12],l[27],l[28]])
                else:
                    pass
        for list in d:
            for l in list :
                if l[2] == '2' and l[7] !='' :
                    l[1] = '生产总成'
                    if len(l) > 34:
                        newData.append([l[1],l[7],l[10],l[11],l[31],str(l[17])+str(l[29]),'',str(l[33])+str(l[34]),'',l[12],l[27],l[28]])
                elif l[3] == '3' and l[7] !='':
                    l[1] = '.1'
                    if len(l) > 34:
                        newData.append([l[1],l[7],l[10],l[11],l[31],str(l[17])+str(l[29]),'',str(l[33])+str(l[34]),'',l[12],l[27],l[28]])
                else:
                    pass
        for list in d:
            for l in list :
                if l[3] == '3' and l[7] !='' :
                    l[1] = '生产总成'
                    if len(l) > 34:
                        newData.append([l[1],l[7],l[10],l[11],l[31],str(l[17])+str(l[29]),'',str(l[33])+str(l[34]),'',l[12],l[27],l[28]])
                elif l[4] == '4' and l[7] !='':
                    l[1] = '.1'
                    if len(l) > 34:
                        newData.append([l[1],l[7],l[10],l[11],l[31],str(l[17])+str(l[29]),'',str(l[33])+str(l[34]),'',l[12],l[27],l[28]])
                else:
                    pass

        autoData = []
        for d in newData:
            if '总成' in d[0]:
                if d not in autoData:
                    autoData.append(d)
                else:
                    pass

        return newData,autoData
