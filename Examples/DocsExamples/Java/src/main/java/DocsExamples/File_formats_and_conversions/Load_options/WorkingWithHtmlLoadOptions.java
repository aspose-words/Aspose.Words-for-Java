package DocsExamples.File_formats_and_conversions.Load_options;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.HtmlLoadOptions;
import com.aspose.words.HtmlControlType;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;

@Test
public class WorkingWithHtmlLoadOptions extends DocsExamplesBase
{
    @Test
    public void preferredControlType() throws Exception {
        //ExStart:LoadHtmlElementsWithPreferredControlType
        final String HTML = "\r\n                <html>\r\n                    <select name='ComboBox' size='1'>\r\n                        <option value='val1'>item1</option>\r\n                        <option value='val2'></option>                        \r\n                    </select>\r\n                </html>\r\n            ";

        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        {
            loadOptions.setPreferredControlType(HtmlControlType.STRUCTURED_DOCUMENT_TAG);
        }

        Document doc = new Document(new ByteArrayInputStream(HTML.getBytes(StandardCharsets.UTF_8)), loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithHtmlLoadOptions.PreferredControlType.docx", SaveFormat.DOCX);
        //ExEnd:LoadHtmlElementsWithPreferredControlType
    }
}
