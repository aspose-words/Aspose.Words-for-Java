package com.aspose.words.examples.asposefeatures.documents.protectdocuments;

import com.aspose.words.Document;
import com.aspose.words.ProtectionType;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class AsposeProtectDoc
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeProtectDoc.class);

        Document doc = new Document(dataDir + "document.doc");
        doc.protect(ProtectionType.READ_ONLY);
        // doc.protect(ProtectionType.ALLOW_ONLY_COMMENTS);
        // doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS);
        // doc.protect(ProtectionType.ALLOW_ONLY_REVISIONS);

        doc.save(dataDir + "AsposeProtect.doc", SaveFormat.DOC);
    }
}
