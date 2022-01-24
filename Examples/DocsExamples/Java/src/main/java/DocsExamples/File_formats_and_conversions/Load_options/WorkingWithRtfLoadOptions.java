package DocsExamples.File_formats_and_conversions.Load_options;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.RtfLoadOptions;
import com.aspose.words.Document;

@Test
public class WorkingWithRtfLoadOptions extends DocsExamplesBase {
    @Test
    public void recognizeUtf8Text() throws Exception {
        //ExStart:RecognizeUtf8Text
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        {
            loadOptions.setRecognizeUtf8Text(true);
        }

        Document doc = new Document(getMyDir() + "UTF-8 characters.rtf", loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithRtfLoadOptions.RecognizeUtf8Text.rtf");
        //ExEnd:RecognizeUtf8Text
    }
}
