package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class OpenTypeFeatures {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		// ExStart:OpenTypeFeatures
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(OpenTypeFeatures.class);

        // Open a document
        Document doc = new Document(dataDir + "OpenType.Document.docx");

        // When text shaper factory is set, layout starts to use OpenType features.
        // An Instance property returns static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory
        doc.getLayoutOptions().setTextShaperFactory(com.aspose.words.shaping.harfbuzz.HarfBuzzTextShaperFactory.getInstance());

        // Render the document to PDF format
        doc.save(dataDir + "OpenType.Document.pdf");
        // ExEnd:OpenTypeFeatures
        System.out.println("\nRendered the document with OpenType Features using HarfBuzz shaping.");

	}

}
