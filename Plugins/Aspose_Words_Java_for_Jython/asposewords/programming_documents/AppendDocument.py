from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import ImportFormatMode
from com.aspose.words import SectionStart

class AppendDocument:

    def __init__(self):
        self.dataDir = Settings.dataDir + 'programming_documents/'
        
        self.simple_append_document()
        self.keep_source_formatting()
        self.use_destination_styles()
        self.join_continuous()
        self.join_new_page()
        self.restart_page_numbering()
        self.link_headers_footers()
        self.unlink_headers_footers()
        self.remove_source_headers_footers()
        self.different_page_setup()
    
    def simple_append_document(self):
        """
            Shows how to append a document to the end of another document using no additional options.
        """
        
        dstDoc = Document(self.dataDir + "TestFile.Destination.doc")
        srcDoc = Document(self.dataDir + "TestFile.Source.doc")

        # Append the source document to the destination document using no extra options.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dstDoc.save(self.dataDir + "TestFile.SimpleAppendDocument Out.docx")
        
    def keep_source_formatting(self):
        """
            Shows how to append a document to the end of another document using no additional options.
        """
        
        dstDoc = Document(self.dataDir + "TestFile.Destination.doc")
        srcDoc = Document(self.dataDir + "TestFile.Source.doc")

        # Append the source document to the destination document using no extra options.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)

        dstDoc.save(self.dataDir + "TestFile.KeepSourceFormatting Out.docx")  
        
    def use_destination_styles(self):
        """
            Shows how to append a document to another document using the formatting of the destination document.
        """
        
        # Load the documents to join.
        dstDoc = Document(self.dataDir + "TestFile.Destination.doc")
        srcDoc = Document(self.dataDir + "TestFile.Source.doc")

        # Append the source document using the styles of the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES)

        # Save the joined document to disk.
        dstDoc.save(self.dataDir + "TestFile.UseDestinationStyles Out.doc")
        
    def join_continuous(self):
        """
            Shows how to append a document to another document so the content flows continuously.
        """
        
        dstDoc = Document(self.dataDir + "TestFile.Destination.doc")
        srcDoc = Document(self.dataDir + "TestFile.Source.doc")

        # Make the document appear straight after the destination documents content.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS)

        # Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)
        
        dstDoc.save(self.dataDir + "TestFile.JoinContinuous Out.doc")
        
    def join_new_page(self):
        """
            Shows how to append a document to another document so it starts on a new page.
        """
        
        dstDoc = Document(self.dataDir + "TestFile.Destination.doc")
        srcDoc = Document(self.dataDir + "TestFile.Source.doc")

        # Set the appended document to start on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE)

        # Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)
        
        dstDoc.save(self.dataDir + "TestFile.JoinNewPage Out.doc")
        
    def restart_page_numbering(self):
        """
            Shows how to append a document to another document with page numbering restarted.
        """
        
        dstDoc = Document(self.dataDir + "TestFile.Destination.doc")
        srcDoc = Document(self.dataDir + "TestFile.Source.doc")

        # Set the appended document to appear on the next page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE)
        # Restart the page numbering for the document to be appended.
        srcDoc.getFirstSection().getPageSetup().setRestartPageNumbering(1)

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)
        
        dstDoc.save(self.dataDir + "TestFile.RestartPageNumbering Out.doc")
        
    def link_headers_footers(self):
        """
            Shows how to append a document to another document and continue headers and footers from the destination document.
        """
        
        dstDoc = Document(self.dataDir + "TestFile.Destination.doc")
        srcDoc = Document(self.dataDir + "TestFile.Source.doc")

        # Set the appended document to appear on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE)

        # Link the headers and footers in the source document to the previous section.
        # This will override any headers or footers already found in the source document.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(1)

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)
        
        dstDoc.save(self.dataDir + "TestFile.LinkHeadersFooters Out.doc")
        
    def unlink_headers_footers(self):
        """
            Shows how to append a document to another document so headers and footers do not continue from the destination document.
        """
        
        dstDoc = Document(self.dataDir + "TestFile.Destination.doc")
        srcDoc = Document(self.dataDir + "TestFile.Source.doc")

        # Even a document with no headers or footers can still have the LinkToPrevious setting set to True.
        # Unlink the headers and footers in the source document to stop this from continuing the headers and footers
        # from the destination document.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(0)

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)
        
        dstDoc.save(self.dataDir + "TestFile.UnlinkHeadersFooters Out.doc")
        
    def remove_source_headers_footers(self):
        """
            Shows how to remove headers and footers from a document before appending it to another document.
        """
        
        dstDoc = Document(self.dataDir + "TestFile.Destination.doc")
        srcDoc = Document(self.dataDir + "TestFile.Source.doc")

        # Remove the headers and footers from each of the sections in the source document.
        for section in srcDoc.getSections().toArray():
            section.clearHeadersFooters()

        # Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting
        # for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination
        # document. This should set to false to avoid this behaviour.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(0)

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)
        
        dstDoc.save(self.dataDir + "TestFile.RemoveSourceHeadersFooters Out.doc")
        
    def different_page_setup(self):
        """
            Shows how to append a document to another document continuously which has different page settings.
        """
        
        dstDoc = Document(self.dataDir + "TestFile.Destination.doc")
        srcDoc = Document(self.dataDir + "TestFile.Source.doc")

        # Set the source document to continue straight after the end of the destination document.
        # If some page setup settings are different then this may not work and the source document will appear
        # on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS)

        # To ensure this does not happen when the source document has different page setup settings make sure the
        # settings are identical between the last section of the destination document.
        # If there are further continuous sections that follow on in the source document then this will need to be
        # repeated for those sections as well.
        srcDoc.getFirstSection().getPageSetup().setPageWidth(dstDoc.getLastSection().getPageSetup().getPageWidth())
        srcDoc.getFirstSection().getPageSetup().setPageHeight(dstDoc.getLastSection().getPageSetup().getPageHeight())
        srcDoc.getFirstSection().getPageSetup().setOrientation(dstDoc.getLastSection().getPageSetup().getOrientation())

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING)
        
        dstDoc.save(self.dataDir + "TestFile.DifferentPageSetup Out.doc")

if __name__ == '__main__':        
    AppendDocument()