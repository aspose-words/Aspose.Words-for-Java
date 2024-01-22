package DocsExamples.LINQ_Reporting_Engine;

import DocsExamples.DocsExamplesBase;
import TestData.Common;
import TestData.TestBuilders.ColorItemTestBuilder;
import TestData.TestClasses.ClientTestClass;
import TestData.TestClasses.ColorItemTestClass;
import TestData.TestClasses.ManagerTestClass;
import TestData.TestClasses.SenderTestClass;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ReportingEngine;

import java.awt.*;
import java.util.ArrayList;

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

        SenderTestClass sender = new SenderTestClass();
        sender.setName("LINQ Reporting Engine");
        sender.setMessage("Hello World");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, sender, "sender");

        doc.save(getArtifactsDir() + "ReportingEngine.HelloWorld.docx");
        //ExEnd:HelloWorld
    }

    @Test
    public void singleRow() throws Exception
    {
        //ExStart:SingleRow
        Document doc = new Document(getMyDir() + "Reporting engine template - Table row (Java).docx");

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(ManagerTestClass.class);
        engine.buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.SingleRow.docx");
        //ExEnd:SingleRow
    }

    @Test
    public void commonMasterDetail() throws Exception
    {
        //ExStart:CommonMasterDetail
        Document doc = new Document(getMyDir() + "Reporting engine template - Common master detail (Java).docx");

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(ManagerTestClass.class);
        engine.buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.CommonMasterDetail.docx");
        //ExEnd:CommonMasterDetail
    }

    @Test
    public void conditionalBlocks() throws Exception
    {
        //ExStart:ConditionalBlocks
        Document doc = new Document(getMyDir() + "Reporting engine template - Table row conditional blocks (Java).docx");

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(ClientTestClass.class);
        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.ConditionalBlock.docx");
        //ExEnd:ConditionalBlocks
    }

    @Test
    public void settingBackgroundColor() throws Exception
    {
        //ExStart:SettingBackgroundColor
        Document doc = new Document(getMyDir() + "Reporting engine template - Background color (Java).docx");

        ArrayList<ColorItemTestClass> colors = new ArrayList<>();
        colors.add(new ColorItemTestBuilder().withColor("Black", Color.BLACK).build());
        colors.add(new ColorItemTestBuilder().withColor("Red", new Color((255), (0), (0))).build());
        colors.add(new ColorItemTestBuilder().withColor("Empty", null).build());

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(ColorItemTestClass.class);
        engine.buildReport(doc, colors, "Colors");

        doc.save(getArtifactsDir() + "ReportingEngine.BackColor.docx");
        //ExEnd:SettingBackgroundColor
    }
}
