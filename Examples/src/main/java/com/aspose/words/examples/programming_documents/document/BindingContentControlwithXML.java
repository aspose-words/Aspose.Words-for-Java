package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.util.UUID;

/**
 * Created by Home on 5/29/2017.
 */
public class BindingContentControlwithXML {
    public static void main(String[] args) throws Exception {

        //ExStart:BindingContentControlwithXML
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(BindingContentControlwithXML.class);

        Document doc = new Document();

        CustomXmlPart xmlPart = doc.getCustomXmlParts().add(UUID.fromString("38400000-8cf0-11bd-b23e-10b96e4ef00d").toString(), "<root><text>Hello, World!</text></root>");

        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PLAIN_TEXT, MarkupLevel.BLOCK);
        doc.getFirstSection().getBody().appendChild(sdt);

        sdt.getXmlMapping().setMapping(xmlPart, "/root[1]/text[1]", "");

        dataDir = dataDir + "BindSDTtoCustomXmlPart_out.doc";

        // Save the document to disk.
        doc.save(dataDir);
        //ExEnd:BindingContentControlwithXML
    }
}
