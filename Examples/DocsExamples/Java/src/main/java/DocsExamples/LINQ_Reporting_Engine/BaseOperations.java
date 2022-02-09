package DocsExamples.LINQ_Reporting_Engine;

import DocsExamples.DocsExamplesBase;
import DocsExamples.LINQ_Reporting_Engine.Helpers.Common;
import DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects.*;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ReportingEngine;
import java.util.ArrayList;
import java.awt.Color;

@Test
public class BaseOperations extends DocsExamplesBase
{
    @Test
    public void helloWorld() throws Exception
    {
        //ExStart:HelloWorld
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.write("<<[sender.getName()]>> says: <<[sender.getMessage()]>>");

        Sender sender = new Sender(); { sender.setName("LINQ Reporting Engine"); sender.setMessage("Hello World"); }

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, sender, "sender");

        doc.save(getArtifactsDir() + "ReportingEngine.HelloWorld.docx");
        //ExEnd:HelloWorld
    }

    @Test
    public void singleRow() throws Exception
    {
        //ExStart:SingleRow
        Document doc = new Document(getMyDir() + "Reporting engine template - Table row.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(Manager.class);
        engine.buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.SingleRow.docx");
        //ExEnd:SingleRow
    }

    @Test
    public void commonMasterDetail() throws Exception
    {
        //ExStart:CommonMasterDetail
        Document doc = new Document(getMyDir() + "Reporting engine template - Common master detail.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(Manager.class);
        engine.buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.CommonMasterDetail.docx");
        //ExEnd:CommonMasterDetail
    }

    @Test
    public void conditionalBlocks() throws Exception
    {
        //ExStart:ConditionalBlocks
        Document doc = new Document(getMyDir() + "Reporting engine template - Table row conditional blocks.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(Client.class);
        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.ConditionalBlock.docx");
        //ExEnd:ConditionalBlocks
    }

    @Test
    public void settingBackgroundColor() throws Exception
    {
        //ExStart:SettingBackgroundColor
        Document doc = new Document(getMyDir() + "Reporting engine template - Background color.docx");

        ArrayList<BackgroundColor> colors = new ArrayList<>();
        {
            colors.add(new BackgroundColor()); {
            colors.get(0).setName("Black"); colors.get(0).setColor(Color.BLACK);}
            colors.add(new BackgroundColor()); {
            colors.get(1).setName("Red"); colors.get(1).setColor(new Color((255), (0), (0)));}
            colors.add(new BackgroundColor()); {
            colors.get(2).setName("Empty"); colors.get(2).setColor(null);}
        }

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(BackgroundColor.class);
        engine.buildReport(doc, colors, "Colors");

        doc.save(getArtifactsDir() + "ReportingEngine.BackColor.docx");
        //ExEnd:SettingBackgroundColor
    }
}
