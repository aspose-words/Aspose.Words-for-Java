from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import NodeType
from com.aspose.words import ControlChar

class RemoveBreaks:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        # Open the document.
        doc = Document(dataDir + "TestRemoveBreaks.doc")

        # Remove the page and section breaks from the document.
        # In Aspose.Words section breaks are represented as separate Section nodes in the document.
        # To remove these separate sections the sections are combined.
        self.remove_page_breaks(doc)
        self.remove_section_breaks(doc)

        # Save the document.
        doc.save(dataDir + "TestRemoveBreaks Out.doc")
    
    def remove_page_breaks(self, doc):
        # Retrieve all paragraphs in the document.
        paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, True)

        paragraphs_count = paragraphs.getCount()

        i = 0
        while(i < paragraphs_count) :
            para = paragraphs.get(i)
            if (para.getParagraphFormat().getPageBreakBefore()):
                para.getParagraphFormat().setPageBreakBefore(False)

            runs = para.getRuns().toArray()

            for run in runs:
                if (run.getText() in ControlChar.PAGE_BREAK):
                    run.setText(run.getText().replace(ControlChar.PAGE_BREAK, ""))

            i = i + 1
        
    def remove_section_breaks(self, doc):
        # Loop through all sections starting from the section that precedes the last one
        # and moving to the first section.
        i = doc.getSections().getCount() - 2
        while ( i >= 0 ):
            # Copy the content of the current section to the beginning of the last section.
            doc.getLastSection().prependContent(doc.getSections().get(i))
            # Remove the copied section.
            doc.getSections().get(i).remove()
            i = i - 1

if __name__ == '__main__':        
    RemoveBreaks()