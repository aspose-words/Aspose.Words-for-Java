/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package programmingwithdocuments.workingwithdocument.removebreaks.java;

import java.io.File;
import java.net.URI;

import com.aspose.words.*;


public class RemoveBreaks
{
    public static void main(String[] args) throws Exception
    {
        // The sample infrastructure.
        String dataDir = "src/programmingwithdocuments/workingwithdocument/removebreaks/data/";

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");

        // Remove the page and section breaks from the document.
        // In Aspose.Words section breaks are represented as separate Section nodes in the document.
        // To remove these separate sections the sections are combined.
        removePageBreaks(doc);
        removeSectionBreaks(doc);

        // Save the document.
        doc.save(dataDir + "TestFile Out.doc");
    }

    //ExStart
    //ExFor:ControlChar.PageBreak
    //ExId:RemoveBreaks_Pages
    //ExSummary:Removes all page breaks from the document.
    private static void removePageBreaks(Document doc) throws Exception
    {
        // Retrieve all paragraphs in the document.
        NodeCollection paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, true);

        // Iterate through all paragraphs
        for (Paragraph para : (Iterable<Paragraph>) paragraphs)
        {
            // If the paragraph has a page break before set then clear it.
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
    //ExEnd


    //ExStart
    //ExId:RemoveBreaks_Sections
    //ExSummary:Combines all sections in the document into one.
    private static void removeSectionBreaks(Document doc) throws Exception
    {
        // Loop through all sections starting from the section that precedes the last one
        // and moving to the first section.
        for (int i = doc.getSections().getCount() - 2; i >= 0; i--)
        {
            // Copy the content of the current section to the beginning of the last section.
            doc.getLastSection().prependContent(doc.getSections().get(i));
            // Remove the copied section.
            doc.getSections().get(i).remove();
        }
    }
    //ExEnd
}