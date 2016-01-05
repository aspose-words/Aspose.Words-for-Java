from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import Shape
from com.aspose.words import ShapeType
from com.aspose.words import HeaderFooterType
from com.aspose.words import HeaderFooter
from com.aspose.words import RelativeHorizontalPosition
from com.aspose.words import RelativeVerticalPosition
from com.aspose.words import WrapType
from com.aspose.words import VerticalAlignment
from com.aspose.words import HorizontalAlignment
from com.aspose.words import Paragraph

from java.awt import Color

class AddWatermark:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        doc = Document(dataDir + "TestFile.doc")
        
        self.insert_watermark_text(doc, "CONFIDENTIAL")
        
        doc.save(dataDir + "Watermark.doc")

        print "Watermark added to the document successfully."
    
    def insert_watermark_text(self, doc, watermarkText):
        # Create a watermark shape. This will be a WordArt shape.
        # You are free to try other shape types as watermarks.
        watermark = Shape(doc, ShapeType.TEXT_PLAIN_TEXT)

        # Set up the text of the watermark.
        watermark.getTextPath().setText(watermarkText)
        watermark.getTextPath().setFontFamily("Arial")
        watermark.setWidth(500)
        watermark.setHeight(100)
        # Text will be directed from the bottom-left to the top-right corner.
        watermark.setRotation(-40)
        # Remove the following two lines if you need a solid black text.
        watermark.getFill().setColor(Color.GRAY) # Try LightGray to get more Word-style watermark
        watermark.setStrokeColor(Color.GRAY) # Try LightGray to get more Word-style watermark

        # Place the watermark in the page center.
        watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE)
        watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE)
        watermark.setWrapType(WrapType.NONE)
        watermark.setVerticalAlignment(VerticalAlignment.CENTER)
        watermark.setHorizontalAlignment(HorizontalAlignment.CENTER)

        # Create a new paragraph and append the watermark to this paragraph.
        watermarkPara = Paragraph(doc)
        watermarkPara.appendChild(watermark)

        # Insert the watermark into all headers of each document section.
        for sect in doc.getSections() :
            # There could be up to three different headers in each section, since we want
            # the watermark to appear on all pages, insert into all headers.
            self.insert_watermark_into_header(watermarkPara, sect, HeaderFooterType.HEADER_PRIMARY)
            self.insert_watermark_into_header(watermarkPara, sect, HeaderFooterType.HEADER_FIRST)
            self.insert_watermark_into_header(watermarkPara, sect, HeaderFooterType.HEADER_EVEN)

    def insert_watermark_into_header(self, watermarkPara, sect, headerType):

        header = sect.getHeadersFooters().getByHeaderFooterType(headerType)

        if (header is None):
            # There is no header of the specified type in the current section, create it.
            header = HeaderFooter(sect.getDocument(), headerType)
            sect.getHeadersFooters().add(header)

        # Insert a clone of the watermark into the header.
        header.appendChild(watermarkPara.deepClone(True))

if __name__ == '__main__':        
    AddWatermark()