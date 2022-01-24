package DocsExamples.Programming_with_Documents.Contents_Management;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.words.ControlChar;
import com.aspose.words.Section;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterType;
import java.util.ArrayList;
import com.aspose.words.FieldStart;
import com.aspose.words.Node;
import com.aspose.words.FieldType;
import com.aspose.words.FieldEnd;


class RemoveContent extends DocsExamplesBase
{
    @Test
    public void removePageBreaks() throws Exception
    {
        //ExStart:OpenFromFile
        Document doc = new Document(getMyDir() + "Document.docx");
        //ExEnd:OpenFromFile

        // In Aspose.Words section breaks are represented as separate Section nodes in the document.
        // To remove these separate sections, the sections are combined.
        removePageBreaks(doc);
        removeSectionBreaks(doc);

        doc.save(getArtifactsDir() + "RemoveContent.RemovePageBreaks.docx");
    }

    //ExStart:RemovePageBreaks
    private void removePageBreaks(Document doc)
    {
        NodeCollection paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, true);

        for (Paragraph para : (Iterable<Paragraph>) paragraphs)
        {
            // If the paragraph has a page break before the set, then clear it.
            if (para.getParagraphFormat().getPageBreakBefore())
                para.getParagraphFormat().setPageBreakBefore(false);

            // Check all runs in the paragraph for page breaks and remove them.
            for (Run run : (Iterable<Run>) para.getRuns())
            {
                if (run.getText().contains(ControlChar.PAGE_BREAK))
                    run.setText(run.getText().replace(ControlChar.PAGE_BREAK, ""));
            }
        }
    }
    //ExEnd:RemovePageBreaks

    //ExStart:RemoveSectionBreaks
    private void removeSectionBreaks(Document doc)
    {
        // Loop through all sections starting from the section that precedes the last one and moving to the first section.
        for (int i = doc.getSections().getCount() - 2; i >= 0; i--)
        {
            // Copy the content of the current section to the beginning of the last section.
            doc.getLastSection().prependContent(doc.getSections().get(i));
            // Remove the copied section.
            doc.getSections().get(i).remove();
        }
    }
    //ExEnd:RemoveSectionBreaks

    @Test
    public void removeFooters() throws Exception
    {
        //ExStart:RemoveFooters
        Document doc = new Document(getMyDir() + "Header and footer types.docx");

        for (Section section : (Iterable<Section>) doc)
        {
            // Up to three different footers are possible in a section (for first, even and odd pages)
            // we check and delete all of them.
            HeaderFooter footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_FIRST);
            footer?.Remove();

            // Primary footer is the footer used for odd pages.
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
            footer?.Remove();

            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_EVEN);
            footer?.Remove();
        }

        doc.save(getArtifactsDir() + "RemoveContent.RemoveFooters.docx");
        //ExEnd:RemoveFooters
    }

    @Test
    //ExStart:RemoveTOCFromDocument
    public void removeToc() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table of contents.docx");

        // Remove the first table of contents from the document.
        removeTableOfContents(doc, 0);

        doc.save(getArtifactsDir() + "RemoveContent.RemoveToc.doc");
    }

    /// <summary>
    /// Removes the specified table of contents field from the document.
    /// </summary>
    /// <param name="doc">The document to remove the field from.</param>
    /// <param name="index">The zero-based index of the TOC to remove.</param>
    public void removeTableOfContents(Document doc, int index)
    {
        // Store the FieldStart nodes of TOC fields in the document for quick access.
        ArrayList<FieldStart> fieldStarts = new ArrayList<FieldStart>();
        // This is a list to store the nodes found inside the specified TOC. They will be removed at the end of this method.
        ArrayList<Node> nodeList = new ArrayList<Node>();

        for (FieldStart start : (Iterable<FieldStart>) doc.getChildNodes(NodeType.FIELD_START, true))
        {
            if (start.getFieldType() == FieldType.FIELD_TOC)
            {
                fieldStarts.add(start);
            }
        }

        // Ensure the TOC specified by the passed index exists.
        if (index > fieldStarts.size() - 1)
            throw new IllegalArgumentException("Specified argument was out of the range of valid values.\r\nParameter name: " + "TOC index is out of range");

        boolean isRemoving = true;
        
        Node currentNode = fieldStarts.get(index);
        while (isRemoving)
        {
            // It is safer to store these nodes and delete them all at once later.
            nodeList.add(currentNode);
            currentNode = currentNode.nextPreOrder(doc);

            // Once we encounter a FieldEnd node of type FieldTOC,
            // we know we are at the end of the current TOC and stop here.
            if (currentNode.getNodeType() == NodeType.FIELD_END)
            {
                FieldEnd fieldEnd = (FieldEnd) currentNode;
                if (fieldEnd.getFieldType() == FieldType.FIELD_TOC)
                    isRemoving = false;
            }
        }

        for (Node node : nodeList)
        {
            node.remove();
        }
    }
    //ExEnd:RemoveTOCFromDocument
}
