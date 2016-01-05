from asposewords import Settings
from java.io import File
from com.aspose.words import LoadFormat
from com.aspose.words import FileFormatUtil

class CheckFormat:

    def __init__(self):
        dataDir = Settings.dataDir + 'loading_saving/CheckFormat/'
            
        filesList  = File(dataDir).listFiles()
        for file in filesList:
            if file.isDirectory():
                continue
            
            nameOnly  = file.getName()
            print nameOnly
            
            fileName = file.getPath()
            print fileName

            info = FileFormatUtil.detectFileFormat(fileName)
            if (info.getLoadFormat() == LoadFormat.DOC):
                print "Microsoft Word 97-2003 document."
            elif (info.getLoadFormat() == LoadFormat.DOT):
                print "Microsoft Word 97-2003 template."
            elif (info.getLoadFormat() == LoadFormat.DOCX):
                print "Office Open XML WordprocessingML Macro-Free Document."
            elif (info.getLoadFormat() == LoadFormat.DOCM):
                print "Office Open XML WordprocessingML Macro-Enabled Document."
            elif (info.getLoadFormat() == LoadFormat.DOTX):
                print "Office Open XML WordprocessingML Macro-Free Template."
            elif (info.getLoadFormat() == LoadFormat.DOTM):
                print "Office Open XML WordprocessingML Macro-Enabled Template."
            elif (info.getLoadFormat() == LoadFormat.FLAT_OPC):
                print "Flat OPC document."
            elif (info.getLoadFormat() == LoadFormat.RTF):
                print "RTF format."
            elif (info.getLoadFormat() == LoadFormat.WORD_ML):
                print "Microsoft Word 2003 WordprocessingML format."
            elif (info.getLoadFormat() == LoadFormat.HTML):
                print "HTML format."
            elif (info.getLoadFormat() == LoadFormat.MHTML):
                print "MHTML (Web archive) format."
            elif (info.getLoadFormat() == LoadFormat.ODT):
                print "OpenDocument Text."
            elif (info.getLoadFormat() == LoadFormat.OTT):
                print "OpenDocument Text Template."
            elif (info.getLoadFormat() == LoadFormat.DOC_PRE_WORD_97):
                print "MS Word 6 or Word 95 format."
            elif (info.getLoadFormat() == LoadFormat.UNKNOWN):
                print "Unknown format."
            else :
                print "Unknown format."

        print "Process Completed Successfully"    

if __name__ == '__main__':
    CheckFormat()