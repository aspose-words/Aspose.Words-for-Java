package DocsExamples.Programming_with_documents.Working_with_document;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.text.MessageFormat;

@Test
public class WorkingWithWebExtension extends DocsExamplesBase
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

        for (TaskPane taskPaneInfo : doc.getWebExtensionTaskPanes())
        {
            WebExtensionReference reference = taskPaneInfo.getWebExtension().getReference();
            System.out.println(MessageFormat.format("Provider: \"{0}\", version: \"{1}\", catalog identifier: \"{2}\";", reference.getStore(), reference.getVersion(), reference.getId()));
        }
        //ExEnd:GetListOfAddins
    }
}
