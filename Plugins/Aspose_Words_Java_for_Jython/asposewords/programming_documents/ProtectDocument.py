from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import ProtectionType
from com.aspose.words import SaveFormat

class ProtectDocument:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        doc = Document(dataDir + "document.doc")

        doc.protect(ProtectionType.READ_ONLY)
        #doc.protect(ProtectionType.ALLOW_ONLY_COMMENTS)
        #doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS)
        #doc.protect(ProtectionType.ALLOW_ONLY_REVISIONS)

        doc.save(dataDir + "ProtectDocument.doc", SaveFormat.DOC)

        print "Done."

if __name__ == '__main__':
    ProtectDocument()