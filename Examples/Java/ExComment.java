//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;


public class ExComment extends ExBase
{
    @Test
    public void acceptAllRevisions() throws Exception
    {
        //ExStart
        //ExFor:Document.AcceptAllRevisions
        //ExId:AcceptAllRevisions
        //ExSummary:Shows how to accept all tracking changes in the document.
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.acceptAllRevisions();
        //ExEnd
    }
}

