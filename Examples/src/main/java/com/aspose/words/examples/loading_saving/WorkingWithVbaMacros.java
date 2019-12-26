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
        CloneVbaProject(dataDir);
        CloneVbaModule(dataDir);
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
        Document doc = new Document(dataDir + "VbaProject_out.docm");

        for (VbaModule module : doc.getVbaProject().getModules()) {
            System.out.println(module.getSourceCode());
        }
        
        doc.save(dataDir + "VbaProject_out.docm");
        //ExEnd: ReadVbaMacros
        System.out.println("\nReading of VBA Macro is successful.");
    }
	
	public static void ModifyVbaMacros(String dataDir) throws Exception
    {
        //ExStart:ModifyVbaMacros
        Document doc = new Document(dataDir + "VbaProject_out.docm");
        VbaProject project = doc.getVbaProject();

        String newSourceCode = "Test change source code";

        // Choose a module, and set a new source code.
        project.getModules().get(0).setSourceCode(newSourceCode);
        //ExEnd:ModifyVbaMacros
        System.out.println("\nModified successfully.");
    }

	public static void CloneVbaProject(String dataDir) throws Exception
    {
        //ExStart:CloneVbaProject
        Document doc = new Document(dataDir + "VbaProject_out.docm");
        VbaProject project = doc.getVbaProject();

        Document destDoc = new Document();

        // Clone the whole project.
        destDoc.setVbaProject(doc.getVbaProject().deepClone());

        destDoc.save(dataDir + "output.docm");
        //ExEnd:CloneVbaProject
        System.out.println("\nCloned Vba Project successfully.\nFile saved at " + dataDir);
    }

	public static void CloneVbaModule(String dataDir) throws Exception
    {
        //ExStart:CloneVbaModule
		Document doc = new Document(dataDir + "VbaProject_out.docm");
        VbaProject project = doc.getVbaProject();

        Document destDoc = new Document();

        destDoc.setVbaProject(new VbaProject());

        // Clone a single module.
        VbaModule copyModule = doc.getVbaProject().getModules().get("Module1").deepClone();
        destDoc.getVbaProject().getModules().add(copyModule);

        destDoc.save(dataDir + "output.docm");
        //ExEnd:CloneVbaModule
        System.out.println("\nCloned Vba Module successfully.\nFile saved at " + dataDir);
    }
}