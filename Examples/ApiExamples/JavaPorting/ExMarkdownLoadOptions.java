// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.MarkdownLoadOptions;
import com.aspose.words.Document;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.words.Underline;


@Test
public class ExMarkdownLoadOptions extends ApiExampleBase
{
    @Test
    public void preserveEmptyLines() throws Exception
    {
        //ExStart:PreserveEmptyLines
        //GistId:a775441ecb396eea917a2717cb9e8f8f
        //ExFor:MarkdownLoadOptions
        //ExFor:MarkdownLoadOptions.#ctor
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

    @Test
    public void importUnderlineFormatting() throws Exception
    {
        //ExStart:ImportUnderlineFormatting
        //GistId:e06aa7a168b57907a5598e823a22bf0a
        //ExFor:MarkdownLoadOptions.ImportUnderlineFormatting
        //ExSummary:Shows how to recognize plus characters "++" as underline text formatting.
        MemoryStream stream = new MemoryStream(Encoding.getASCII().getBytes("++12 and B++"));
        try /*JAVA: was using*/
        {
            MarkdownLoadOptions loadOptions = new MarkdownLoadOptions(); { loadOptions.setImportUnderlineFormatting(true); }
            Document doc = new Document(stream, loadOptions);

            Paragraph para = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);
            Assert.assertEquals(Underline.SINGLE, para.getRuns().get(0).getFont().getUnderline());

            loadOptions = new MarkdownLoadOptions(); { loadOptions.setImportUnderlineFormatting(false); }
            doc = new Document(stream, loadOptions);

            para = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);
            Assert.assertEquals(Underline.NONE, para.getRuns().get(0).getFont().getUnderline());
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd:ImportUnderlineFormatting
    }
}

