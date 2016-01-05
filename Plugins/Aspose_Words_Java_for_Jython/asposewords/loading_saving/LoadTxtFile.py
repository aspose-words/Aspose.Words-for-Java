from asposewords import Settings
from com.aspose.words import Document

class LoadTxtFile:

    def __init__(self):
        dataDir = Settings.dataDir + 'loading_saving/'
        
        # The encoding of the text file is automatically detected.
        doc = Document(dataDir + "LoadTxt.txt")

        # Save as any Aspose.Words supported format, such as DOCX.
        doc.save(dataDir + "LoadTxt_Out.docx")

        print "Process Completed Successfully"

if __name__ == '__main__':
    LoadTxtFile()