package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.VbaModule;
import com.aspose.words.examples.Utils;

public class WorkingWithVbaMacros {

	public static void main(String[] args) throws Exception {
		//ExStart: ReadVbaMacros
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithVbaMacros.class);

        Document doc = new Document(dataDir + "Document.dot");
        
        for (VbaModule module : doc.getVbaProject().getModules()) {
                System.out.println(module.getSourceCode());
        }
        //ExEnd: ReadVbaMacros
	}
}