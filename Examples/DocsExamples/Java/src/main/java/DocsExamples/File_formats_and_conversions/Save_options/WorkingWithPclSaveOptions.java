package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.PclSaveOptions;
import com.aspose.words.SaveFormat;

@Test
public class WorkingWithPclSaveOptions extends DocsExamplesBase
{
    @Test
    public void rasterizeTransformedElements() throws Exception
    {
        //ExStart:RasterizeTransformedElements
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PclSaveOptions saveOptions = new PclSaveOptions();
        {
            saveOptions.setSaveFormat(SaveFormat.PCL); saveOptions.setRasterizeTransformedElements(false);
        }

        doc.save(getArtifactsDir() + "WorkingWithPclSaveOptions.RasterizeTransformedElements.pcl", saveOptions);
        //ExEnd:RasterizeTransformedElements
    }
}
