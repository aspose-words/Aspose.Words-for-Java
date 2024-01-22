package DocsExamples.LINQ_Reporting_Engine;

import DocsExamples.DocsExamplesBase;
import TestData.Common;
import TestData.TestClasses.ClientTestClass;
import TestData.TestClasses.ManagerTestClass;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.aspose.words.DocumentBuilder;

@Test
public class Lists extends DocsExamplesBase
{
    @Test
    public void createBulletedList() throws Exception
    {
        //ExStart:BulletedList
        Document doc = new Document(getMyDir() + "Reporting engine template - Bulleted list (Java).docx");

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(ClientTestClass.class);
        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.CreateBulletedList.docx");
        //ExEnd:BulletedList
    }

    @Test
    public void inParagraphList() throws Exception
    {
        //ExStart:InParagraphList
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.write("<<foreach [ClientTestClass in clients]>><<[indexOf() !=0 ? ”, ”:  ””]>><<[getName()]>><</foreach>>");
        
        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(ClientTestClass.class);

        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.InParagraphList.docx");
        //ExEnd:InParagraphList
    }

    @Test
    public void inTableList() throws Exception
    {
        //ExStart:InTableList
        Document doc = new Document(getMyDir() + "Reporting engine template - Contextual object member access (Java).docx");

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(ManagerTestClass.class);
        engine.buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.InTableList.docx");
        //ExEnd:InTableList
    }

    @Test
    public void multicoloredNumberedList() throws Exception
    {
        //ExStart:MulticoloredNumberedList
        Document doc = new Document(getMyDir() + "Reporting engine template - Multicolored numbered list (Java).docx");

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(ClientTestClass.class);
        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.MulticoloredNumberedList.doc");
        //ExEnd:MulticoloredNumberedList
    }

    @Test
    public void numberedList() throws Exception
    {
        //ExStart:NumberedList
        Document doc = new Document(getMyDir() + "Reporting engine template - Numbered list (Java).docx");

        ReportingEngine engine = new ReportingEngine();
        engine.getKnownTypes().add(ClientTestClass.class);
        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.NumberedList.docx");
        //ExEnd:NumberedList
    }
}
