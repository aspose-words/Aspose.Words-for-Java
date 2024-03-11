package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.MarkdownLoadOptions;
import com.aspose.words.Document;
import org.testng.Assert;


class ExMarkdownLoadOptions extends ApiExampleBase
{
    @Test
    public void preserveEmptyLines() throws Exception
    {
        //ExStart:PreserveEmptyLines
        //GistId:a775441ecb396eea917a2717cb9e8f8f
        //ExFor:MarkdownLoadOptions
        //ExFor:MarkdownLoadOptions.PreserveEmptyLines
        //ExSummary:Shows how to preserve empty line while load a document.
        String mdText = $"{Environment.NewLine}Line1{Environment.NewLine}{Environment.NewLine}Line2{Environment.NewLine}{Environment.NewLine}";
        MemoryStream stream = new MemoryStream(Encoding.getUTF8().getBytes(mdText));
        try /*JAVA: was using*/
        {
            MarkdownLoadOptions loadOptions = new MarkdownLoadOptions(); { loadOptions.setPreserveEmptyLines(true); }
            Document doc = new Document(stream, loadOptions);

            Assert.assertEquals("\rLine1\r\rLine2\r\f", doc.getText());
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd:PreserveEmptyLines
    }
}

