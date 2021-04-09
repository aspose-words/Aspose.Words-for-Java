// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.ComHelper;
import com.aspose.words.Document;
import org.testng.Assert;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;


@Test
public class ExComHelper extends ApiExampleBase
{
    @Test
    public void comHelper() throws Exception
    {
        //ExStart
        //ExFor:ComHelper
        //ExFor:ComHelper.#ctor
        //ExFor:ComHelper.Open(Stream)
        //ExFor:ComHelper.Open(String)
        //ExSummary:Shows how to open documents using the ComHelper class.
        // The ComHelper class allows us to load documents from within COM clients.
        ComHelper comHelper = new ComHelper();

        // 1 -  Using a local system filename:
        Document doc = comHelper.open(getMyDir() + "Document.docx");

        Assert.assertEquals("Hello World!\r\rHello Word!\r\r\rHello World!", doc.getText().trim());

        // 2 -  From a stream:
        FileStream stream = new FileStream(getMyDir() + "Document.docx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            doc = comHelper.open(stream);

            Assert.assertEquals("Hello World!\r\rHello Word!\r\r\rHello World!", doc.getText().trim());
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd
    }
}

