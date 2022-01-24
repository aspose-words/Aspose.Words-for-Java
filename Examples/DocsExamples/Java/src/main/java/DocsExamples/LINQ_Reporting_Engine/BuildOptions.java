package DocsExamples.LINQ_Reporting_Engine;

import DocsExamples.DocsExamplesBase;
import DocsExamples.LINQ_Reporting_Engine.Helpers.Common;
import DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects.Manager;
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
        Document doc = new Document(getMyDir() + "Reporting engine template - Remove empty paragraphs.docx");

        ReportingEngine engine = new ReportingEngine(); { engine.setOptions(ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS); }
        engine.getKnownTypes().add(Manager.class);
        engine.buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.RemoveEmptyParagraphs.docx");
        //ExEnd:RemoveEmptyParagraphs
    }
}
