package DocsExamples.LINQ_Reporting_Engine;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects.Sender;
import com.aspose.words.ReportingEngine;
import java.util.ArrayList;
import DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects.BackgroundColor;
import java.awt.Color;
import com.aspose.ms.System.Drawing.msColor;


public class BaseOperations extends DocsExamplesBase
{
    @Test
    public void helloWorld() throws Exception
    {
        //ExStart:HelloWorld
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.write("<<[sender.Name]>> says: <<[sender.Message]>>");

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
        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.ConditionalBlock.docx");
        //ExEnd:ConditionalBlocks
    }

    @Test
    public void settingBackgroundColor() throws Exception
    {
        //ExStart:SettingBackgroundColor
        Document doc = new Document(getMyDir() + "Reporting engine template - Background color.docx");

        ArrayList<BackgroundColor> colors = new ArrayList<BackgroundColor>();
        {
            colors.add(new BackgroundColor()); {((BackgroundColor)colors.get(0)).setName("Black"); ((BackgroundColor)colors.get(0)).setColor(Color.BLACK);}
            colors.add(new BackgroundColor()); {((BackgroundColor)colors.get(1)).setName("Red"); ((BackgroundColor)colors.get(1)).setColor(new Color((255), (0), (0)));}
            colors.add(new BackgroundColor()); {((BackgroundColor)colors.get(2)).setName("Empty"); ((BackgroundColor)colors.get(2)).setColor(msColor.Empty);}
        }

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, colors, "Colors");

        doc.save(getArtifactsDir() + "ReportingEngine.BackColor.docx");
        //ExEnd:SettingBackgroundColor
    }
}
