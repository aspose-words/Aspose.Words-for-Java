__author__ = 'fahadadeel'
import jpype
import shutil

class CheckFormat:

    def __init__(self,dataDir):

        self.dataDir = dataDir
        self.Document = jpype.JClass("com.aspose.words.Document")
        self.File = jpype.JClass("java.io.File")
        self.LoadFormat = jpype.JClass("com.aspose.words.LoadFormat")
        self.infoObj = jpype.JClass('com.aspose.words.FileFormatUtil')

    def main(self):
        
        supportedDir = self.dataDir + '/OutSupported/'
        fileObj = self.File(self.dataDir)
        filesList  = fileObj.listFiles()
        for file in filesList:
            if file.isDirectory:
                continue

            nameOnly  = file.getName()
            print (nameOnly)

            fileName = file.getPath()
            print (fileName)

            info = self.infoObj.detectFileFormat(fileName)
            if (info.getLoadFormat() == self.LoadFormat.DOC):
                print ("Microsoft Word 97-2003 document.")
            elif (info.getLoadFormat() == self.LoadFormat.DOT):
                print ("Microsoft Word 97-2003 template.")
            elif (info.getLoadFormat() == self.LoadFormat.DOCX):
                print ("Office Open XML WordprocessingML Macro-Free Document.")
            elif (info.getLoadFormat() == self.LoadFormat.DOCM):
                print ("Office Open XML WordprocessingML Macro-Enabled Document.")
            elif (info.getLoadFormat() == self.LoadFormat.DOTX):
                print ("Office Open XML WordprocessingML Macro-Free Template.")
            elif (info.getLoadFormat() == self.LoadFormat.DOTM):
                print ("Office Open XML WordprocessingML Macro-Enabled Template.")
            elif (info.getLoadFormat() == self.LoadFormat.FLAT_OPC):
                print ("Flat OPC document.")
            elif (info.getLoadFormat() == self.LoadFormat.RTF):
                print ("RTF format.")
            elif (info.getLoadFormat() == self.LoadFormat.WORD_ML):
                print ("Microsoft Word 2003 WordprocessingML format.")
            elif (info.getLoadFormat() == self.LoadFormat.HTML):
                print ("HTML format.")
            elif (info.getLoadFormat() == self.LoadFormat.MHTML):
                print ("MHTML (Web archive) format.")
            elif (info.getLoadFormat() == self.LoadFormat.ODT):
                print ("OpenDocument Text.")
            elif (info.getLoadFormat() == self.LoadFormat.OTT):
                print ("OpenDocument Text Template.")
            elif (info.getLoadFormat() == self.LoadFormat.DOC_PRE_WORD_97):
                print ("MS Word 6 or Word 95 format.")
            elif (info.getLoadFormat() == self.LoadFormat.UNKNOWN):
                print ("Unknown format.")
            else :
                print ("Unknown format.")

            destFileObj = self.File(supportedDir + nameOnly)
            destFile = destFileObj.getPath()
            shutil.copy(fileName, destFile)