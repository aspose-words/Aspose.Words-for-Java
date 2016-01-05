from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import SaveFormat

class ConvertToFormats:

    def __init__(self):
        dataDir = Settings.dataDir + 'loading_saving/'
            
        # Load the document from disk.
        doc = Document(dataDir + "document.doc")
        
        doc.save(dataDir + "Aspose_DocToHTML.html",SaveFormat.HTML) # Save the document in HTML format.
        doc.save(dataDir + "Aspose_DocToPDF.pdf",SaveFormat.PDF) # Save the document in PDF format.
        doc.save(dataDir + "Aspose_DocToTxt.txt",SaveFormat.TEXT) # Save the document in TXT format.
        doc.save(dataDir + "Aspose_DocToJPG.jpg",SaveFormat.JPEG) # Save the document in JPEG format.

        print "Doc file converted in specified formats"

if __name__ == '__main__':
    ConvertToFormats()