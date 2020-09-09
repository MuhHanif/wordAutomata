from docx import Document as dc
import csv

class editor(object):
    """read, edit, and write document"""

    def __init__(self):

        #set global variable
        pass

    def entryTemplateFieldContent(self, dir, documentName, csvData):

        #read document
        doc = dc(dir)

        #count paragraph
        pGraphCount = len(doc.paragraphs)

        #get this data from csv and loop
        with open(csvData) as dataCsv:

            #convert csv to list
            readCsv = list(csv.reader(dataCsv, delimiter=","))

            pass

        print(readCsv)

        #loop through csv data
        for x in range(len(readCsv)):

            #read specific data from csv
            replaceFieldFound = readCsv[x][0]
            textReplacement = readCsv[x][1]
            print(replaceFieldFound)
            print(textReplacement)

            #read all paragraph
            for paragraph in range(0, pGraphCount):

                #get each paragraph
                pGraph = doc.paragraphs[paragraph]

                #count run
                runCount = len(pGraph.runs)

                #read all
                for run in range(0, runCount):

                    #get each run
                    runData = pGraph.runs[run]

                    #replace field with desired text
                    if replaceFieldFound in runData.text:

                        #notification
                        print("paragraph",paragraph,"run",run,"[",replaceFieldFound,"terganti dengan", textReplacement,"]")
                        runData.text = runData.text.replace(replaceFieldFound, textReplacement)

                        pass

                        pass

            pass
        doc.save(documentName)

        pass

while True:

    try:

        pass
        #document directory
        data = input("masukkan directory word: ")
        replace = input("masukan directory csv pengganti: ")
        output = input("nama file output: ")

        file = editor().entryTemplateFieldContent(data, output, replace)

        pass

    except:

        print("error salah entry")
        pass
