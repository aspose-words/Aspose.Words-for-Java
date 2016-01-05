from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import ImageSaveOptions
from com.aspose.words import SaveFormat
from com.aspose.words import TiffCompression

class SaveAsMultipageTiff:

    def __init__(self):
        dataDir = Settings.dataDir + 'rendering_printing/'
            
        # Open the document.
        doc = Document(dataDir + "TestFile.doc")

        # Save the document as multipage TIFF.
        doc.save(dataDir + "TestFile Out.tiff")
        
        # Create an ImageSaveOptions object to pass to the Save method
        options = ImageSaveOptions(SaveFormat.TIFF)
        options.setPageIndex(0)
        options.setPageCount(2)
        options.setTiffCompression(TiffCompression.CCITT_4)
        options.setResolution(160)

        doc.save(dataDir + "TestFileWithOptions Out.tiff", options)

        "Document saved as multi page TIFF successfully."

if __name__ == '__main__':
    SaveAsMultipageTiff()