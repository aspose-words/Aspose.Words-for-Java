package com.aspose.words.examples.programming_documents.WebExtension;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.TaskPane;
import com.aspose.words.TaskPaneDockState;
import com.aspose.words.WebExtensionBinding;
import com.aspose.words.WebExtensionBindingType;
import com.aspose.words.WebExtensionProperty;
import com.aspose.words.WebExtensionStoreType;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.quickstart.AppendDocuments;

public class WorkingWithWebExtension {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub

		// The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithWebExtension.class);
		UsingWebExtensionTaskPanes(dataDir);
	}
	
	public static void UsingWebExtensionTaskPanes(String dataDir) throws Exception
    {
        // ExStart:UsingWebExtensionTaskPanes
        Document doc = new Document();

        TaskPane taskPane = new TaskPane();
        doc.getWebExtensionTaskPanes().add(taskPane);

        taskPane.setDockState(TaskPaneDockState.RIGHT);
        taskPane.isVisible(true);
        taskPane.setWidth(300);

        taskPane.getWebExtension().getReference().setId("wa102923726");
        taskPane.getWebExtension().getReference().setVersion("1.0.0.0");
        taskPane.getWebExtension().getReference().setStoreType(WebExtensionStoreType.OMEX);
        taskPane.getWebExtension().getReference().setStore("th-TH");
        taskPane.getWebExtension().getProperties().add(new WebExtensionProperty("mailchimpCampaign", "mailchimpCampaign"));
        taskPane.getWebExtension().getBindings().add(new WebExtensionBinding("UnnamedBinding_0_1506535429545", WebExtensionBindingType.TEXT, "194740422"));
        
        doc.save(dataDir + "output.docx", SaveFormat.DOCX);
        // ExEnd:UsingWebExtensionTaskPanes 
        System.out.println("\nThe file is saved successfully at " + dataDir);
    }

}
