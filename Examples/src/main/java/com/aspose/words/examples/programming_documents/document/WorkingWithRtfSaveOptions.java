package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.RtfSaveOptions;
import com.aspose.words.examples.Utils;

public class WorkingWithRtfSaveOptions {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertElements.class);
        
		SavingImagesAsWmf(dataDir);
	}
	
	public static void SavingImagesAsWmf(String dataDir) throws Exception
    {
        // ExStart:SavingImagesAsWmf 
        String fileName = "TestFile.doc";
        Document doc = new Document(dataDir + fileName);

        RtfSaveOptions saveOpts = new RtfSaveOptions();
        saveOpts.setSaveImagesAsWmf(true);

        doc.save(dataDir + "output.rtf", saveOpts);
        //ExEnd:SavingImagesAsWmf
    }

}
