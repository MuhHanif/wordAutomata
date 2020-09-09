from docx import Document as dc



class editor(object):
    """read, edit, and write document"""

    def __init__(self):

        #set global variable
        pass



    def entryTemplateFieldContent(self, dir):


        #read document
        doc = dc(dir)

        #count paragraph
        pGraphCount = len(doc.paragraphs)

        #read all paragraph
        for paragraph in range(0, pGraphCount):

            #count run
            runCount = len(doc.paragraphs[paragraph].runs)

            #read all
            for run in range(0, runCount):

                #print paragraph to cmd
                print(paragraph," ",run," ","[" + doc.paragraphs[paragraph].runs[run].text + "]")

                pass

            pass

        pass

#document directory
data = input("masukkan directory word: ")

file = editor().entryTemplateFieldContent(data)
