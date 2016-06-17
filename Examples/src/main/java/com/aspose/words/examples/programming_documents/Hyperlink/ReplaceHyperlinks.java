package com.aspose.words.examples.programming_documents.Hyperlink;

import com.aspose.words.Document;
import com.aspose.words.Field;
import com.aspose.words.FieldHyperlink;
import com.aspose.words.FieldType;
import com.aspose.words.examples.Utils;


public class ReplaceHyperlinks {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ReplaceHyperlinks.class);
        String NewUrl = "http://www.aspose.com";
        String NewName = "Aspose - The .NET & Java Component Publisher";

        // Open the document.
        Document doc = new Document(dataDir + "ReplaceHyperlinks.doc");
        for (Field field : doc.getRange().getFields()) {
            if (field.getType() == FieldType.FIELD_HYPERLINK) {
                FieldHyperlink hyperlink = (FieldHyperlink) field;

                // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                if (hyperlink.getSubAddress() != null)
                    continue;

                hyperlink.setAddress(NewUrl);
                hyperlink.setResult(NewName);
            }
            doc.save(dataDir + "output.doc");
            //ExEnd:1
        }
    }

}