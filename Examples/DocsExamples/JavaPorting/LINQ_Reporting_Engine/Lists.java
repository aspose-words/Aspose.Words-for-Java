package DocsExamples.LINQ_Reporting_Engine;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.aspose.words.DocumentBuilder;


public class Lists extends DocsExamplesBase
{
    @Test
    public void createBulletedList() throws Exception
    {
        //ExStart:BulletedList
        Document doc = new Document(getMyDir() + "Reporting engine template - Bulleted list.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.CreateBulletedList.docx");
        //ExEnd:BulletedList
    }

    @Test
    public void commonList() throws Exception
    {
        //ExStart:CommonList
        Document doc = new Document(getMyDir() + "Reporting engine template - Common master detail.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.CommonList.docx");
        //ExEnd:CommonList
    }

    @Test
    public void inParagraphList() throws Exception
    {
        //ExStart:InParagraphList
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.write("<<foreach [in clients]>><<[IndexOf() !=0 ? ”, ”:  ””]>><<[Name]>><</foreach>>");
        
        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.InParagraphList.docx");
        //ExEnd:InParagraphList
    }

    @Test
    public void inTableList() throws Exception
    {
        //ExStart:InTableList
        Document doc = new Document(getMyDir() + "Reporting engine template - Contextual object member access.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.InTableList.docx");
        //ExEnd:InTableList
    }

    @Test
    public void multicoloredNumberedList() throws Exception
    {
        //ExStart:MulticoloredNumberedList
        Document doc = new Document(getMyDir() + "Reporting engine template - Multicolored numbered list.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.MulticoloredNumberedList.doc");
        //ExEnd:MulticoloredNumberedList
    }

    @Test
    public void numberedList() throws Exception
    {
        //ExStart:NumberedList
        Document doc = new Document(getMyDir() + "Reporting engine template - Numbered list.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getClients(), "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.NumberedList.docx");
        //ExEnd:NumberedList
    }
}
