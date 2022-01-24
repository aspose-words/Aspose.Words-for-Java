package DocsExamples.File_Formats_and_Conversions.Load_Options;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.HtmlLoadOptions;
import com.aspose.words.HtmlControlType;
import com.aspose.words.Document;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.SaveFormat;


public class WorkingWithHtmlLoadOptions extends DocsExamplesBase
{
    @Test
    public void preferredControlType() throws Exception
    {
        //ExStart:LoadHtmlElementsWithPreferredControlType
        final String HTML = "\r\n                <html>\r\n                    <select name='ComboBox' size='1'>\r\n                        <option value='val1'>item1</option>\r\n                        <option value='val2'></option>                        \r\n                    </select>\r\n                </html>\r\n            ";

        HtmlLoadOptions loadOptions = new HtmlLoadOptions(); { loadOptions.setPreferredControlType(HtmlControlType.STRUCTURED_DOCUMENT_TAG); }

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(HTML)), loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithHtmlLoadOptions.PreferredControlType.docx", SaveFormat.DOCX);
        //ExEnd:LoadHtmlElementsWithPreferredControlType
    }
}
