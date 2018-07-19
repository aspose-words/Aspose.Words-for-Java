package com.aspose.words.examples.programming_documents.document.properties;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.loading_saving.WorkingWithTxt;

public class DocProperties {
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getDataDir(DocProperties.class);

        removePersonalInformation(dataDir);
    }

    public static void removePersonalInformation(String dataDir) throws Exception
    {
        // ExStart:RemovePersonalInformation
        Document doc = new Document(dataDir + "Properties.doc");
        doc.setRemovePersonalInformation(true);

        dataDir = dataDir + "RemovePersonalInformation_out.docx";
        doc.save(dataDir);
        // ExEnd:RemovePersonalInformation
        System.out.println("\nPersonal information has been removed from document successfully.\nFile saved at " + dataDir);
    }
}
