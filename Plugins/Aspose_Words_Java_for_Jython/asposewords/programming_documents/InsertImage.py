from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import DocumentBuilder
from com.aspose.words import RelativeHorizontalPosition
from com.aspose.words import RelativeVerticalPosition
from com.aspose.words import WrapType

class InsertImage:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        doc = Document()

        builder = DocumentBuilder(doc)

        builder.insertImage(dataDir + "background.jpg");
        builder.insertImage(dataDir + "background.jpg",
                RelativeHorizontalPosition.MARGIN,
                100,
                RelativeVerticalPosition.MARGIN,
                200,
                200,
                100,
                WrapType.SQUARE)

        doc.save(dataDir + "InsertImage.docx")

        print "Done."

if __name__ == '__main__':
    InsertImage()