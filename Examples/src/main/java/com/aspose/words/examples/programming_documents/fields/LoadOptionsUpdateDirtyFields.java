package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.LoadOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 9/18/2017.
 */
public class LoadOptionsUpdateDirtyFields {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ConvertFieldsInDocument.class);
        // ExStart:LoadOptionsUpdateDirtyFields
        LoadOptions lo = new LoadOptions();

        //Update the fields with the dirty attribute
        lo.setUpdateDirtyFields(true);

        //Load the Word document
        Document doc = new Document(dataDir + "input.docx", lo);

        //Save the document into DOCX
        doc.save(dataDir + "output.docx", SaveFormat.DOCX);
        // ExEnd:LoadOptionsUpdateDirtyFields
        System.out.println("Updated the fields with the dirty attribute successfully");
    }
}
