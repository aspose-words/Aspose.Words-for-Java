package com.aspose.words.examples.asposefeatures.workingwithdocument.movingcursor;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Node;
import com.aspose.words.Paragraph;
import com.aspose.words.examples.Utils;

public class AsposeMovingCursor
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeMovingCursor.class);

        Document doc = new Document(dataDir + "document.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        //Shows how to access the current node in a document builder.
        Node curNode = builder.getCurrentNode();
        Paragraph curParagraph = builder.getCurrentParagraph();

        // Shows how to move a cursor position to a specified node.
        builder.moveTo(doc.getFirstSection().getBody().getLastParagraph());

        // Shows how to move a cursor position to the beginning or end of a document.
        builder.moveToDocumentEnd();
        builder.writeln("This is the end of the document.");

        builder.moveToDocumentStart();
        builder.writeln("This is the beginning of the document.");

        doc.save(dataDir + "AsposeMovingCursor.doc");

        System.out.println("Done.");
    }
}


