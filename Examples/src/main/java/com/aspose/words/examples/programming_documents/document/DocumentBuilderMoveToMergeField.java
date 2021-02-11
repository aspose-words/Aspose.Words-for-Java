package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Field;
import org.testng.Assert;

public class DocumentBuilderMoveToMergeField {
    public static void main(String[] args) throws Exception {
        //ExStart:DocumentBuilderMoveToMergeField
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field using the DocumentBuilder and add a run of text after it.
        Field field = builder.insertField("MERGEFIELD field");
        builder.write(" Text after the field.");

        // The builder's cursor is currently at end of the document.
        Assert.assertNull(builder.getCurrentNode());
        // We can move the builder to a field like this, placing the cursor at immediately after the field.
        builder.moveToField(field, true);

        // Note that the cursor is at a place past the FieldEnd node of the field, meaning that we are not actually inside the field.
        // If we wish to move the DocumentBuilder to inside a field,
        // we will need to move it to a field's FieldStart or FieldSeparator node using the DocumentBuilder.MoveTo() method.
        Assert.assertEquals(field.getEnd(), builder.getCurrentNode().getPreviousSibling());
        builder.write(" Text immediately after the field.");
        //ExEnd:DocumentBuilderMoveToMergeField
    }
}