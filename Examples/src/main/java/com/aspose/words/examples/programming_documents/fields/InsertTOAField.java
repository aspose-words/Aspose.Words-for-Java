package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 5/29/2017.
 */
public class InsertTOAField {
    public static void main(String[] args) throws Exception {
        //ExStart:InsertTOAField
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertTOAField.class);

        Document doc = new Document(dataDir + "in.doc");

        // Get paragraph you want to append this TOA field to
        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(1);

        // We want to insert TA and TOA fields like this:
        // { TA \c 1 \l "Value 0" }
        // { TOA \c 1 }

        // Create instance of FieldA class and lets build the above field code
        FieldTA fieldTA = (FieldTA) para.appendField(FieldType.FIELD_TOA_ENTRY, false);
        fieldTA.setEntryCategory("1");
        fieldTA.setLongCitation("Value 0");

        doc.getFirstSection().getBody().appendChild(para);

        para = new Paragraph(doc);

        // Create instance of FieldToa class
        FieldToa fieldToa = (FieldToa) para.appendField(FieldType.FIELD_TOA, false);
        fieldToa.setEntryCategory("1");
        doc.getFirstSection().getBody().appendChild(para);

        // Finally update this TOA field
        fieldToa.update();

        dataDir = dataDir + "InsertTOAFieldWithoutDocumentBuilder_out.doc";
        doc.save(dataDir);
        //ExEnd:InsertTOAField
    }
}
