package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.text.MessageFormat;

@Test
class ExMarkdownLoadOptions extends ApiExampleBase
{
    @Test
    public void preserveEmptyLines() throws Exception
    {
        //ExStart:PreserveEmptyLines
        //GistId:9c17d666c47318436785490829a3984f
        //ExFor:MarkdownLoadOptions
        //ExFor:MarkdownLoadOptions.#ctor
        //ExFor:MarkdownLoadOptions.PreserveEmptyLines
        //ExSummary:Shows how to preserve empty line while load a document.
        String mdText = MessageFormat.format("{0}Line1{0}{0}Line2{0}{0}", System.lineSeparator());

        MarkdownLoadOptions loadOptions = new MarkdownLoadOptions();
        loadOptions.setPreserveEmptyLines(true);
        Document doc = new Document(new ByteArrayInputStream(mdText.getBytes()), loadOptions);

        Assert.assertEquals("\rLine1\r\rLine2\r\f", doc.getText());
        //ExEnd:PreserveEmptyLines
    }

    @Test
    public void importUnderlineFormatting() throws Exception
    {
        //ExStart:ImportUnderlineFormatting
        //GistId:6280fd6c1c1854468bea095ec2af902b
        //ExFor:MarkdownLoadOptions.ImportUnderlineFormatting
        //ExSummary:Shows how to recognize plus characters "++" as underline text formatting.
        try (ByteArrayInputStream stream = new ByteArrayInputStream("++12 and B++".getBytes(StandardCharsets.US_ASCII)))
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
        //ExEnd:ImportUnderlineFormatting
    }
}

