package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.shaping.harfbuzz.HarfBuzzTextShaperFactory;


class EnableOpenTypeFeatures extends DocsExamplesBase
{
    @Test
    public void openTypeFeatures() throws Exception
    {
        //ExStart:OpenTypeFeatures
        Document doc = new Document(getMyDir() + "OpenType text shaping.docx");

        // When we set the text shaper factory, the layout starts to use OpenType features.
        // An Instance property returns BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory.
        doc.getLayoutOptions().setTextShaperFactory(com.aspose.words.shaping.harfbuzz.HarfBuzzTextShaperFactory.getInstance());

        doc.save(getArtifactsDir() + "WorkingWithHarfBuzz.OpenTypeFeatures.pdf");
        //ExEnd:OpenTypeFeatures
    }
}
