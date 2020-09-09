from docx import Document as dc
import pandas as pd


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

        replacementFile = pd.read_csv(csvData, header=None)

        print(replacementFile)

        for x in range(len(replacementFile.iloc[:,0])):

            replaceFieldFound = replacementFile.iloc[x, 0]
            textReplacement = replacementFile.iloc[x, 1]
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
                        print("[",replaceFieldFound,"terganti dengan", textReplacement,"]")
                        runData.text = runData.text.replace(replaceFieldFound, textReplacement)
                        print(runData.text)

                        #print paragraph to cmd
                        print(paragraph," ",run," ","[" + doc.paragraphs[paragraph].runs[run].text + "]")

                        pass

                        pass

            pass
        doc.save(documentName)

        pass

#document directory
data = input("masukkan directory word: ")

file = editor().entryTemplateFieldContent(data, "result.docx", "replace.csv")
