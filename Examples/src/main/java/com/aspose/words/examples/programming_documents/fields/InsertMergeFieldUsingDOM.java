package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class InsertMergeFieldUsingDOM {
    public static void main(String[] args) throws Exception {

        //ExStart:InsertMergeFieldUsingDOM
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertMergeFieldUsingDOM.class);

        Document doc = new Document(dataDir + "in.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Get paragraph you want to append this merge field to
        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(1);

        // Move cursor to this paragraph
        builder.moveTo(para);

        // We want to insert a merge field like this:
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }

        // Create instance of FieldMergeField class and lets build the above field code
        FieldMergeField field = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, false);

        // { " MERGEFIELD Test1" }
        field.setFieldName("Test1");

        // { " MERGEFIELD Test1 \\b Test2" }
        field.setTextBefore("Test2");

        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
        field.setTextAfter("Test3");

        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
        field.isMapped(true);

        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
        field.isVerticalFormatting(true);

        // Finally update this merge field
        field.update();

        doc.save(dataDir + "output.docx");
        //ExEnd:InsertMergeFieldUsingDOM


    }
}




