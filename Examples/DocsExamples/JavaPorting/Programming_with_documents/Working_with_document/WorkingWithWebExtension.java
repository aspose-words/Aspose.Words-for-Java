package DocsExamples.Programming_with_Documents.Working_with_Document;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.TaskPane;
import com.aspose.words.TaskPaneDockState;
import com.aspose.words.WebExtensionStoreType;
import com.aspose.words.WebExtensionProperty;
import com.aspose.words.WebExtensionBinding;
import com.aspose.words.WebExtensionBindingType;
import com.aspose.ms.System.msConsole;
import com.aspose.words.WebExtensionReference;


class WorkingWithWebExtension extends DocsExamplesBase
{
    @Test
    public void usingWebExtensionTaskPanes() throws Exception
    {
        //ExStart:UsingWebExtensionTaskPanes
        Document doc = new Document();

        TaskPane taskPane = new TaskPane();
        doc.getWebExtensionTaskPanes().add(taskPane);

        taskPane.setDockState(TaskPaneDockState.RIGHT);
        taskPane.isVisible(true);
        taskPane.setWidth(300.0);

        taskPane.getWebExtension().getReference().setId("wa102923726");
        taskPane.getWebExtension().getReference().setVersion("1.0.0.0");
        taskPane.getWebExtension().getReference().setStoreType(WebExtensionStoreType.OMEX);
        taskPane.getWebExtension().getReference().setStore("th-TH");
        taskPane.getWebExtension().getProperties().add(new WebExtensionProperty("mailchimpCampaign", "mailchimpCampaign"));
        taskPane.getWebExtension().getBindings().add(new WebExtensionBinding("UnnamedBinding_0_1506535429545",
            WebExtensionBindingType.TEXT, "194740422"));

        doc.save(getArtifactsDir() + "WorkingWithWebExtension.UsingWebExtensionTaskPanes.docx");
        //ExEnd:UsingWebExtensionTaskPanes
        
        //ExStart:GetListOfAddins
        doc = new Document(getArtifactsDir() + "WorkingWithWebExtension.UsingWebExtensionTaskPanes.docx");
        
        System.out.println("Task panes sources:\n");

        for (TaskPane taskPaneInfo : (Iterable<TaskPane>) doc.getWebExtensionTaskPanes())
        {
            WebExtensionReference reference = taskPaneInfo.getWebExtension().getReference();
            System.out.println("Provider: \"{reference.Store}\", version: \"{reference.Version}\", catalog identifier: \"{reference.Id}\";");
        }
        //ExEnd:GetListOfAddins
    }
}
