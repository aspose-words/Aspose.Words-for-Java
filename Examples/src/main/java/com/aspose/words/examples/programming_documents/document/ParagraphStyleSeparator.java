package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.examples.Utils;

public class ParagraphStyleSeparator {

    public static void main(String[] args) throws Exception {
        // ExStart: ParagraphStyleSeparator
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ParagraphStyleSeparator.class);

        // Initialize document.
        String fileName = "TestFile.doc";
        Document doc = new Document(dataDir + fileName);

        NodeCollection paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, true);
        for (Paragraph paragraph : (Iterable<Paragraph>) paragraphs) {
            if (paragraph.getBreakIsStyleSeparator()) {
                System.out.println("Separator Found!");
            }
        }
        // ExEnd: ParagraphStyleSeparator
    }
}
