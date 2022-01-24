package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.OdtSaveOptions;
import com.aspose.words.OdtSaveMeasureUnit;

@Test
public class WorkingWithOdtSaveOptions extends DocsExamplesBase
{
    @Test
    public void measureUnit() throws Exception
    {
        //ExStart:MeasureUnit
        Document doc = new Document(getMyDir() + "Document.docx");

        // Open Office uses centimeters when specifying lengths, widths and other measurable formatting
        // and content properties in documents whereas MS Office uses inches.
        OdtSaveOptions saveOptions = new OdtSaveOptions(); { saveOptions.setMeasureUnit(OdtSaveMeasureUnit.INCHES); }

        doc.save(getArtifactsDir() + "WorkingWithOdtSaveOptions.MeasureUnit.odt", saveOptions);
        //ExEnd:MeasureUnit
    }
}
