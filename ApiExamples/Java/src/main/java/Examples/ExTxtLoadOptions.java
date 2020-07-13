package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Document;
import com.aspose.words.DocumentDirection;
import com.aspose.words.TxtLoadOptions;
import org.testng.Assert;
import org.testng.annotations.Test;

@Test
public class ExTxtLoadOptions extends ApiExampleBase {
    @Test
    public void detectDocumentDirection() throws Exception {
        //ExStart
        //ExFor:TxtLoadOptions.DocumentDirection
        //ExSummary:Shows how to detect document direction automatically.
        // Create a LoadOptions object and configure it to detect text direction automatically upon loading
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.setDocumentDirection(DocumentDirection.AUTO);

        // Text like Hebrew/Arabic will be automatically detected as RTL
        Document doc = new Document(getMyDir() + "Hebrew text.txt", loadOptions);

        Assert.assertTrue(doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBidi());

        doc = new Document(getMyDir() + "English text.txt", loadOptions);

        Assert.assertFalse(doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBidi());
        //ExEnd
    }
}
