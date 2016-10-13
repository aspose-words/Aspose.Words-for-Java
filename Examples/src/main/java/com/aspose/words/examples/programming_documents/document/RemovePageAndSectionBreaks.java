package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class RemovePageAndSectionBreaks {
    public static void main(String[] args) throws Exception {
        
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RemovePageAndSectionBreaks.class);

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");

        // Remove the page and section breaks from the document.
        // In Aspose.Words section breaks are represented as separate Section nodes in the document.
        // To remove these separate sections the sections are combined.
        removePageBreaks(doc);
        removeSectionBreaks(doc);

        // Save the document.
        doc.save(dataDir + "TestFile_out.doc");

        System.out.println("Removed breaks from the document successfully.");
    }

    /* ExSummary:Removes all page breaks from the document.*/
    private static void removePageBreaks(Document doc) throws Exception {
        // Retrieve all paragraphs in the document.
        NodeCollection paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, true);

        // Iterate through all paragraphs
        for (Paragraph para : (Iterable<Paragraph>) paragraphs) {
            // If the paragraph has a page break before set then clear it.
            if (para.getParagraphFormat().getPageBreakBefore())
                para.getParagraphFormat().setPageBreakBefore(false);

            // Check all runs in the paragraph for page breaks and remove them.
            for (Run run : para.getRuns()) {
                if (run.getText().contains(ControlChar.PAGE_BREAK))
                    run.setText(run.getText().replace(ControlChar.PAGE_BREAK, ""));
            }
        }
    }

    /* ExSummary:Combines all sections in the document into one.*/
    private static void removeSectionBreaks(Document doc) throws Exception {
        // Loop through all sections starting from the section that precedes the last one
        // and moving to the first section.
        for (int i = doc.getSections().getCount() - 2; i >= 0; i--) {
            // Copy the content of the current section to the beginning of the last section.
            doc.getLastSection().prependContent(doc.getSections().get(i));
            // Remove the copied section.
            doc.getSections().get(i).remove();
        }
    }
}