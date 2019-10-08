package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 5/29/2017.
 */
public class InsertINCLUDETEXT {
    public static void main(String[] args) throws Exception {

        //ExStart:InsertINCLUDETEXT
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertINCLUDETEXT.class);

        Document doc = new Document(dataDir + "in.doc");

        //Get paragraph you want to append this INCLUDETEXT field to
        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(1);

        //We want to insert an INCLUDETEXT field like this:
        //{ INCLUDETEXT  "file path" }

        //Create instance of FieldAsk class and lets build the above field code
        FieldIncludeText fieldincludeText = (FieldIncludeText) para.appendField(FieldType.FIELD_INCLUDE_TEXT, false);
        fieldincludeText.setBookmarkName("bookmark");
        fieldincludeText.setSourceFullName(dataDir + "IncludeText.docx");

        doc.getFirstSection().getBody().appendChild(para);

        //Finally update this INCLUDETEXT field
        fieldincludeText.update();

        dataDir = dataDir + "InsertIncludeFieldWithoutDocumentBuilder_out.doc";

        doc.save(dataDir);
        //ExEnd:InsertINCLUDETEXT

    }
}
