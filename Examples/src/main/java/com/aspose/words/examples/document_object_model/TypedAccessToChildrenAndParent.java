package com.aspose.words.examples.document_object_model;

import com.aspose.words.*;

public class TypedAccessToChildrenAndParent {

    public static void main(String[] args) throws Exception {
        //ExStart:
        Document doc = new Document();

        // Quick typed access to the first child Section node of the Document.
        Section section = doc.getFirstSection();

        // Quick typed access to the Body child node of the Section.
        Body body = section.getBody();

        // Quick typed access to all Table child nodes contained in the Body.
        TableCollection tables = body.getTables();

        for (Table table : tables) {
            // Quick typed access to the first row of the table.
            if (table.getFirstRow() != null)
                table.getFirstRow().remove();

            // Quick typed access to the last row of the table.
            if (table.getLastRow() != null)
                table.getLastRow().remove();
        }
        //ExEnd:
    }
}
