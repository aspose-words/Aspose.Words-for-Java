package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.VbaModule;
import com.aspose.words.VbaProject;
import com.aspose.words.examples.Utils;

public class WorkingWithVbaMacros {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithVbaMacros.class);
        
        ReadVbaMacros(dataDir);
        ModifyVbaMacros(dataDir);
	}
	
	public static void ReadVbaMacros(String dataDir) throws Exception
    {
		//ExStart: ReadVbaMacros
        Document doc = new Document(dataDir + "Document.dot");

        for (VbaModule module : doc.getVbaProject().getModules()) {
            System.out.println(module.getSourceCode());
        }
        //ExEnd: ReadVbaMacros
    }
	
	public static void ModifyVbaMacros(String dataDir) throws Exception
    {
        //ExStart:ModifyVbaMacros
        Document doc = new Document(dataDir + "test.docm");
        VbaProject project = doc.getVbaProject();

        String newSourceCode = "Test change source code";

        // Choose a module, and set a new source code.
        project.getModules().get(0).setSourceCode(newSourceCode);
        //ExEnd:ModifyVbaMacros
    }
}