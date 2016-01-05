from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import DocumentBuilder
from com.aspose.words import HeaderFooterType
from com.aspose.words import ControlChar
from com.aspose.words import SectionStart
from com.aspose.words import TabAlignment
from com.aspose.words import TabLeader
from com.aspose.words import TabStop

from java.io import File
from javax.imageio import ImageIO

class InsertBarcodeOnEachPage:

    def __init__(self):
        self.dataDir = Settings.dataDir + 'programming_documents/'
        
        # Create a blank document.
	doc = Document()
	builder = DocumentBuilder(doc)

	# The number of pages the document should have.
	numPages = 4
	# The document starts with one section, insert the barcode into this existing section.
	self.insert_barcode_into_footer(builder, doc.getFirstSection(), 1, HeaderFooterType.FOOTER_PRIMARY)

	i = 1
	while (i < numPages) :
	    # Clone the first section and add it into the end of the document.
	    cloneSection = doc.getFirstSection().deepClone(False)
	    cloneSection.getPageSetup().setSectionStart(SectionStart.NEW_PAGE)
	    doc.appendChild(cloneSection)

	    # Insert the barcode and other information into the footer of the section.
	    self.insert_barcode_into_footer(builder, cloneSection, i, HeaderFooterType.FOOTER_PRIMARY)
	    i = i + 1

	# Save the document as a PDF to disk. You can also save this directly to a stream.
	doc.save(self.dataDir + "InsertBarcodeOnEachPage.docx")
        
	print "Aspose Barcode Inserted..."
        
    def insert_barcode_into_footer(self, builder, section, pageId, footerType):
	# Move to the footer type in the specific section.
	builder.moveToSection(section.getDocument().indexOf(section))
	builder.moveToHeaderFooter(footerType)

	# Insert the barcode, then move to the next line and insert the ID
	# along with the page number.
	# Use pageId if you need to insert a different barcode on each page. 0
	# = First page, 1 = Second page etc.
	builder.insertImage(ImageIO.read(File(self.dataDir + "barcode.png")))
	builder.writeln()
	builder.write("1234567890")
	builder.insertField("PAGE")

	# Create a right aligned tab at the right margin.
	tabPos = section.getPageSetup().getPageWidth() - section.getPageSetup().getRightMargin() - section.getPageSetup().getLeftMargin()
                
	builder.getCurrentParagraph().getParagraphFormat().getTabStops().add(TabStop(tabPos, TabAlignment.RIGHT, TabLeader.NONE))

	# Move to the right hand side of the page and insert the page and page total.
	builder.write(ControlChar.TAB)
	builder.insertField("PAGE")
	builder.write(" of ")
	builder.insertField("NUMPAGES")

if __name__ == '__main__':
    InsertBarcodeOnEachPage()