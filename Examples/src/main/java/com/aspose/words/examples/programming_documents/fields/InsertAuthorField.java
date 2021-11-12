package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class InsertAuthorField {
    public static void main(String[] args) throws Exception {

        //ExStart:InsertAuthorField
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertAuthorField.class);

        Document doc = new Document(dataDir + "in.doc");

        // Get paragraph you want to append this merge field to
        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(1);

        // We want to insert an AUTHOR field like this:
        // { AUTHOR Test1 }

        // Create instance of FieldAuthor class and lets build the above field code
        FieldAuthor field = (FieldAuthor) para.appendField(FieldType.FIELD_AUTHOR, false);

        // { AUTHOR Test1 }
        field.setAuthorName("Test1");

        // Finally update this AUTHOR field
        field.update();
        doc.save(dataDir + "output.docx");
        //ExEnd:InsertAuthorField
    }
}




