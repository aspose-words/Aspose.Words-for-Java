//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.TxtSaveOptions;
import org.testng.Assert;

public class ExTxtSaveOptions extends ApiExampleBase
{
    @Test
    public void pageBreaks() throws Exception
    {
        //ExStart
        //ExFor:TxtSaveOptions.ForcePageBreaks
        //ExSummary:Shows how to specify whether the page breaks should be preserved during export.
        Document doc = new Document(getMyDir() + "SaveOptions.PageBreaks.docx");

        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.setForcePageBreaks(false);

        doc.save(getArtifactsDir() + "SaveOptions.PageBreaks.False.txt", saveOptions);
        //ExEnd
        Document docFalse = new Document(getArtifactsDir() + "SaveOptions.PageBreaks.False.txt");
        Assert.assertEquals(docFalse.getText(), "Some text before page break\r\rJidqwjidqwojidqwojidqwojidqwojidqwoji\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\rQwdqwdqwdqwdqwdqwdqwd\rQwdqwdqwdqwdqwdqwdqw\r\r\r\r\rqwdqwdqwdqwdqwdqwdqwqwd\r\f");

        saveOptions.setForcePageBreaks(true);
        doc.save(getArtifactsDir() + "SaveOptions.PageBreaks.True.txt", saveOptions);

        Document docTrue = new Document(getArtifactsDir() + "SaveOptions.PageBreaks.True.txt");
        Assert.assertEquals(docTrue.getText(), "Some text before page break\r\f\r\fJidqwjidqwojidqwojidqwojidqwojidqwoji\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\rQwdqwdqwdqwdqwdqwdqwd\rQwdqwdqwdqwdqwdqwdqw\r\r\r\r\f\r\fqwdqwdqwdqwdqwdqwdqwqwd\r\f");
    }
}
