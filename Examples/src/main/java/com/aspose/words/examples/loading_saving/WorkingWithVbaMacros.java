package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.VbaModule;
import com.aspose.words.VbaModuleType;
import com.aspose.words.VbaProject;
import com.aspose.words.examples.Utils;

public class WorkingWithVbaMacros {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithVbaMacros.class);
        
        CreateVbaProject(dataDir);
        ReadVbaMacros(dataDir);
        ModifyVbaMacros(dataDir);
	}
	
	public static void CreateVbaProject(String dataDir) throws Exception
    {
        //ExStart:CreateVbaProject
        Document doc = new Document();

        // Create a new VBA project.
        VbaProject project = new VbaProject();
        project.setName("AsposeProject");
        doc.setVbaProject(project);

        // Create a new module and specify a macro source code.
        VbaModule module = new VbaModule();
        module.setName("AsposeModule");
        module.setType(VbaModuleType.PROCEDURAL_MODULE);
        module.setSourceCode("New source code");

        // Add module to the VBA project.
        doc.getVbaProject().getModules().add(module);

        doc.save(dataDir + "VbaProject_out.docm");
        //ExEnd:CreateVbaProject
        System.out.println("\nDocument saved successfully.\nFile saved at " + dataDir);
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