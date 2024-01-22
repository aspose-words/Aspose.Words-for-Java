package DocsExamples.LINQ_Reporting_Engine;

import DocsExamples.DocsExamplesBase;
import TestData.Common;
import TestData.TestClasses.ManagerTestClass;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.aspose.words.ReportBuildOptions;

@Test
public class BuildOptions extends DocsExamplesBase
{
    @Test
    public void removeEmptyParagraphs() throws Exception
    {
        //ExStart:RemoveEmptyParagraphs
        Document doc = new Document(getMyDir() + "Reporting engine template - Remove empty paragraphs (Java).docx");

        ReportingEngine engine = new ReportingEngine(); { engine.setOptions(ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS); }
        engine.getKnownTypes().add(ManagerTestClass.class);
        engine.buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.RemoveEmptyParagraphs.docx");
        //ExEnd:RemoveEmptyParagraphs
    }
}
