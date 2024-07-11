package DocsExamples.File_Formats_and_Conversions.Save_Options;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.PclSaveOptions;
import com.aspose.words.SaveFormat;


public class WorkingWithPclSaveOptions extends DocsExamplesBase
{
    @Test
    public void rasterizeTransformedElements() throws Exception
    {
        //ExStart:RasterizeTransformedElements
        //GistId:7ee438947078cf070c5bc36a4e45a18c
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PclSaveOptions saveOptions = new PclSaveOptions();
        {
            saveOptions.setSaveFormat(SaveFormat.PCL); saveOptions.setRasterizeTransformedElements(false);
        }

        doc.save(getArtifactsDir() + "WorkingWithPclSaveOptions.RasterizeTransformedElements.pcl", saveOptions);
        //ExEnd:RasterizeTransformedElements
    }
}
