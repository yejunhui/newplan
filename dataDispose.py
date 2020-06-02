
class Dispose:
    def __init__(self):
        pass

    def analyze(self,d):
        #newData
        newData = []
        for list in d:
            for l in list :
                if l[1] == '1':
                    l[1] = '总成'
                    newData.append(l)
                elif l[1] == '2':
                    l[1] = '.1'
                    newData.append(l)
                else:
                    pass
            for l in list :
                if l[1] == '2' :
                    l[1] = '总成'
                    newData.append(l)
                elif l[1] == '3':
                    l[1] = '.1'
                    newData.append(l)
                else:
                    pass
            for l in list :
                if l[1] == '3' :
                    l[1] = '总成'
                    newData.append(l)
                elif l[1] == '4':
                    l[1] = '.1'
                    newData.append(l)
                else:
                    pass

        return newData
