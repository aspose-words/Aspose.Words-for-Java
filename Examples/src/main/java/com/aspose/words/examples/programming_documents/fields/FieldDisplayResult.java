package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.Field;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.joining_appending.AppendwithImportFormatOptions;

public class FieldDisplayResult {

    public static void main(String[] args) throws Exception {
        // ExStart: FieldDisplayResult
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AppendwithImportFormatOptions.class);

        Document document = new Document(dataDir + "Document.docx");
        document.updateFields();

        for (Field field : document.getRange().getFields()) {
            System.out.println(field.getDisplayResult());
        }
        // ExEnd: FieldDisplayResult
    }
}
