from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import SaveFormat
from com.aspose.words import NodeType

class ExtractContentBasedOnStyles:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        # Open the document.
        doc = Document(dataDir + "TestFile.doc")

        # Define style names as they are specified in the Word document.
        PARA_STYLE = "Heading 1"
        RUN_STYLE = "Intense Emphasis"

        # Collect paragraphs with defined styles.
        # Show the number of collected paragraphs and display the text of this paragraphs.
        paragraphs = self.paragraphs_by_style_name(doc, PARA_STYLE)

        print "abc = " + str(paragraphs[0])
        print "Paragraphs with " + PARA_STYLE + " styles " + str(len(paragraphs)) + ":"

        for paragraph in paragraphs :
            print str(paragraph.toString(SaveFormat.TEXT))

        # Collect runs with defined styles.
        # Show the number of collected runs and display the text of this runs.
        runs = self.runs_by_style_name(doc, RUN_STYLE)

        print "Runs with " + RUN_STYLE + " styles " + str(len(runs)) + ":"

        for run in runs :
            print run.getRange().getText()
    
    def paragraphs_by_style_name(self, doc, styleName):

        # Create an array to collect paragraphs of the specified style.
        paragraphsWithStyle = []
        # Get all paragraphs from the document.
        paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, True)
        # Look through all paragraphs to find those with the specified style.

        paragraphs_count = paragraphs.getCount()

        i = 0
        while(i < paragraphs_count) :
            paragraph = paragraphs.get(i)
            if (paragraph.getParagraphFormat().getStyle().getName() == styleName):
                paragraphsWithStyle.append(paragraph)
            i = i + 1

        return paragraphsWithStyle

    def runs_by_style_name(self, doc, styleName):

        # Create an array to collect runs of the specified style.
        runsWithStyle = []

        runs = doc.getChildNodes(NodeType.RUN, True)
        # Look through all runs to find those with the specified style.
        runs = runs.toArray()
        for run in runs :
            if (run.getFont().getStyle().getName() == styleName):
                runsWithStyle.append(run)

        return runsWithStyle

if __name__ == '__main__':        
    ExtractContentBasedOnStyles()