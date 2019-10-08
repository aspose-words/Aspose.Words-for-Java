package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldType;
import com.aspose.words.FieldUnknown;
import com.aspose.words.examples.Utils;

public class InsertFieldNone {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertFieldNone.class);

        // ExStart:InsertFieldNone
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldUnknown field = (FieldUnknown) builder.insertField(FieldType.FIELD_NONE, false);

        dataDir = dataDir + "InsertFieldNone_out.docx";
        doc.save(dataDir);
        // ExEnd:InsertFieldNone
        System.out.println("\nInserted field in the document successfully.\nFile saved at " + dataDir);
    }
}
