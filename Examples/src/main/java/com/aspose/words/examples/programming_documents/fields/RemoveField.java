package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.Field;
import com.aspose.words.examples.Utils;

public class RemoveField {
    public static void main(String[] args) throws Exception {

        //ExStart:RemoveField
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RemoveField.class);
        Document doc = new Document(dataDir + "Field.RemoveField.doc");

        Field field = doc.getRange().getFields().get(0);
        // Calling this method completely removes the field from the document.
        field.remove();

        doc.save(dataDir + "output.docx");
        //ExEnd:RemoveField

        System.out.println("Field removed from the document successfully.");
    }
}




