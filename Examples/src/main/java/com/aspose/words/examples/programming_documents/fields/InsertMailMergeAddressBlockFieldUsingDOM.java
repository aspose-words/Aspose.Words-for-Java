package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class InsertMailMergeAddressBlockFieldUsingDOM {
    public static void main(String[] args) throws Exception {

        //ExStart:InsertMailMergeAddressBlockFieldUsingDOM
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertMailMergeAddressBlockFieldUsingDOM.class);

        Document doc = new Document(dataDir + "in.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Get paragraph you want to append this merge field to
        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(1);

        // Move cursor to this paragraph
        builder.moveTo(para);

        // We want to insert a mail merge address block like this:
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }

        // Create instance of FieldAddressBlock class and lets build the above field code
        FieldAddressBlock field = (FieldAddressBlock) builder.insertField(FieldType.FIELD_ADDRESS_BLOCK, false);

        // { ADDRESSBLOCK \\c 1" }
        field.setIncludeCountryOrRegionName("1");

        // { ADDRESSBLOCK \\c 1 \\d" }
        field.setFormatAddressOnCountryOrRegion(true);

        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 }
        field.setExcludedCountryOrRegionName("Test2");

        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 }
        field.setNameAndAddressFormat("Test3");

        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
        field.setLanguageId("Test4");

        // Finally update this merge field
        field.update();

        doc.save(dataDir + "output.docx");
        //ExEnd:InsertMailMergeAddressBlockFieldUsingDOM
    }
}




