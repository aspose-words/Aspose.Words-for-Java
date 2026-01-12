package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Document;
import com.aspose.words.shaping.harfbuzz.HarfBuzzTextShaperFactory;
import org.testng.annotations.Test;

class EnableOpenTypeFeatures extends DocsExamplesBase {
    @Test
    public void openTypeFeatures() throws Exception {
        //ExStart:OpenTypeFeatures
        //GistId:046c81231cf8e3ba6df901a38b9c861c
        Document doc = new Document(getMyDir() + "OpenType text shaping.docx");

        // When we set the text shaper factory, the layout starts to use OpenType features.
        // An Instance property returns BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory.
        doc.getLayoutOptions().setTextShaperFactory(HarfBuzzTextShaperFactory.getInstance());

        doc.save(getArtifactsDir() + "WorkingWithHarfBuzz.OpenTypeFeatures.pdf");
        //ExEnd:OpenTypeFeatures
    }
}
