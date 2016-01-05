from asposewords import Settings
from com.aspose.words import FileFormatUtil
from java.text import MessageFormat
from java.io import File

class DigitalSignatures:

    def __init__(self):
        dataDir = Settings.dataDir + 'loading_saving/'
            
        # The path to the document which is to be processed.
        filePath = dataDir + "document.doc"
        
        info = FileFormatUtil.detectFileFormat(filePath)
        
        if info.hasDigitalSignature():
            print MessageFormat.format(
                    "Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.",
                    File(doc).getName())
        else:
            print "Document has no digital signature."
            
        print "Process Completed Successfully"

if __name__ == '__main__':
    DigitalSignatures()