from docx import Document as dc
import csv
import xlrd

class editor(object):
    """read, edit, and write document"""

    def __init__(self):

        #set global variable
        pass

    def entryTemplateFieldContent(self, dir, csvData):


        #get this data from csv and loop
        readExcel = xlrd.open_workbook(csvData)
        sheet = readExcel.sheet_by_index(0)


        #loop through csv data

        #iterate through rows
        for x in range(sheet.nrows - 1):
            #read document
            doc = dc(dir)

            #count paragraph
            pGraphCount = len(doc.paragraphs)

            fileName = sheet.cell_value(x + 1,0)
            print(fileName)

            #iterate through columns
            for cols in range(sheet.ncols):

                #read specific data from csv
                replaceFieldFound = sheet.cell_value(0, cols) #this value holds
                textReplacement = sheet.cell_value(x + 1, cols) #this value varies

                #for debug only
                #print(replaceFieldFound)
                #print(textReplacement)

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
                            runData.text = runData.text.replace(replaceFieldFound, str(textReplacement))

                            pass

                        pass

                    pass

                pass

            doc.save(fileName + ".docx")
            print("done! created file as",fileName + ".docx")
            #some bug casues print fist row only !
            pass

        pass

while True:

    try:

        #document directory
        data = input("masukkan directory word: ")
        replace = input("masukan directory excel pengganti: ")

        #break ethernal loop
        if data == "quit":

            break

        #output = input("nama file output: ")

        file = editor().entryTemplateFieldContent(data, replace)

        pass

    except:

        #spit error
        print("error salah entry")

        pass
