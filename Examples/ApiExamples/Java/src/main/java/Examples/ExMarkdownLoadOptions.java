package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Document;
import com.aspose.words.MarkdownLoadOptions;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
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
        //ExFor:MarkdownLoadOptions.PreserveEmptyLines
        //ExSummary:Shows how to preserve empty line while load a document.
        String mdText = MessageFormat.format("{0}Line1{0}{0}Line2{0}{0}", System.lineSeparator());

        MarkdownLoadOptions loadOptions = new MarkdownLoadOptions();
        loadOptions.setPreserveEmptyLines(true);
        Document doc = new Document(new ByteArrayInputStream(mdText.getBytes()), loadOptions);

        Assert.assertEquals("\rLine1\r\rLine2\r\f", doc.getText());
        //ExEnd:PreserveEmptyLines
    }
}

