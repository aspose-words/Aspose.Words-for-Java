package com.aspose.words.examples.programming_documents.Hyperlink;

import com.aspose.words.Document;
import com.aspose.words.Field;
import com.aspose.words.FieldHyperlink;
import com.aspose.words.FieldType;
import com.aspose.words.examples.Utils;

public class ReplaceHyperlinks {

    // The path to the documents directory.
    private static final String dataDir = Utils.getSharedDataDir(ReplaceHyperlinks.class) + "Hyperlink/";

    public static void main(String[] args) throws Exception {

        //ExStart:ReplaceHyperlinks
        String newUrl = "http://www.aspose.com";
        String newName = "Aspose - The .NET & Java Component Publisher";

        // Open the document.
        Document doc = new Document(dataDir + "ReplaceHyperlinks.docx");
        for (Field field : doc.getRange().getFields()) {
            if (field.getType() == FieldType.FIELD_HYPERLINK) {
                FieldHyperlink hyperlink = (FieldHyperlink) field;

                // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                if (hyperlink.getSubAddress() != null)
                    continue;

                hyperlink.setAddress(newUrl);
                hyperlink.setResult(newName);
            }
            doc.save(dataDir + "ReplaceHyperlinks_Out.doc");
            //ExEnd:ReplaceHyperlinks
        }
    }
}