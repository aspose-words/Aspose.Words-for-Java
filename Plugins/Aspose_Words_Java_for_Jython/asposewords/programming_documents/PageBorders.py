from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import DocumentBuilder

class PageBorders:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        doc = Document()
        builder = DocumentBuilder(doc)

        pageSetup = builder.getPageSetup()
        pageSetup.setTopMargin(0.5)
        pageSetup.setBottomMargin(0.5)
        pageSetup.setLeftMargin(0.5)
        pageSetup.setRightMargin(0.5)
        pageSetup.setHeaderDistance(0.2)
        pageSetup.setFooterDistance(0.2)

        doc.save(dataDir + "PageBorders.docx")

        print "Done."

if __name__ == '__main__':
    PageBorders()