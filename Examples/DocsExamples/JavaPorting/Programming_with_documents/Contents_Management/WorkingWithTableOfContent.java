package DocsExamples.Programming_with_Documents.Contents_Management;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.words.TabStop;


class WorkingWithTableOfContent extends DocsExamplesBase
{
    @Test
    public void changeStyleOfTocLevel() throws Exception
    {
        //ExStart:ChangeStyleOfTOCLevel
        Document doc = new Document();
        // Retrieve the style used for the first level of the TOC and change the formatting of the style.
        doc.getStyles().getByStyleIdentifier(StyleIdentifier.TOC_1).getFont().setBold(true);
        //ExEnd:ChangeStyleOfTOCLevel
    }

    @Test
    public void changeTocTabStops() throws Exception
    {
        //ExStart:ChangeTOCTabStops
        Document doc = new Document(getMyDir() + "Table of contents.docx");

        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            // Check if this paragraph is formatted using the TOC result based styles.
            // This is any style between TOC and TOC9.
            if (para.getParagraphFormat().getStyle().getStyleIdentifier() >= StyleIdentifier.TOC_1 &&
                para.getParagraphFormat().getStyle().getStyleIdentifier() <= StyleIdentifier.TOC_9)
            {
                // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                TabStop tab = para.getParagraphFormat().getTabStops().get(0);
                
                // Remove the old tab from the collection.
                para.getParagraphFormat().getTabStops().removeByPosition(tab.getPosition());
                
                // Insert a new tab using the same properties but at a modified position.
                // We could also change the separators used (dots) by passing a different Leader type.
                para.getParagraphFormat().getTabStops().add(tab.getPosition() - 50.0, tab.getAlignment(), tab.getLeader());
            }
        }

        doc.save(getArtifactsDir() + "WorkingWithTableOfContent.ChangeTocTabStops.docx");
        //ExEnd:ChangeTOCTabStops
    }
}
