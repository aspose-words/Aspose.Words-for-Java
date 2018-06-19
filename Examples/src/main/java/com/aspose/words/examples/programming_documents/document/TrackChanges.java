package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.document.properties.AccessingDocumentProperties;

public class TrackChanges {
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getSharedDataDir(AccessingDocumentProperties.class) + "Document/";

        acceptRevisions(dataDir);
        getRevisionTypes(dataDir);
    }

    private static void acceptRevisions(String dataDir) throws Exception {
        // ExStart:TrackChanges
        Document doc = new Document(dataDir + "Document.doc");

        // Start tracking and make some revisions.
        doc.startTrackRevisions("Author");
        doc.getFirstSection().getBody().appendParagraph("Hello world!");

        // Revisions will now show up as normal text in the output document.
        doc.acceptAllRevisions();

        dataDir = dataDir + "Document.AcceptedRevisions_out.doc";
        doc.save(dataDir);
        // ExEnd:AcceptAllRevisions
        System.out.println("\nAll revisions accepted.\nFile saved at " + dataDir);
    }

    private static void getRevisionTypes(String dataDir) throws Exception {
        // ExStart:GetRevisionTypes
        Document doc = new Document(dataDir + "Revisions.docx");

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();
        for (int i = 0; i < paragraphs.getCount(); i++) {
            if (paragraphs.get(i).isMoveFromRevision())
                System.out.println("The paragraph " + i + " has been moved (deleted).");
            if (paragraphs.get(i).isMoveToRevision())
                System.out.println("The paragraph " + i + " has been moved (inserted).");
        }
        // ExEnd:GetRevisionTypes
    }
}
