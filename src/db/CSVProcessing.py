import csv
import os


class CSVProcessing:

    def __init__(self, fileName):
        self.fileName = fileName

    def readCSVtoList(self):
        csvfile = open(self.fileName)
        r = csv.reader(csvfile)
        return (list(r))

    def writeCSVFilter(self, savaPath, saveFileNmae, contextList):
        os.chdir(savaPath)
        csvfile = open(saveFileNmae, 'a+', newline='')
        r = csv.writer(csvfile)
        for row in contextList:
            if "W" in row[3]:
                r.writerow(row)
