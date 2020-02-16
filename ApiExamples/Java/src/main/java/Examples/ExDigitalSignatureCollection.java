package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.DigitalSignature;
import com.aspose.words.DigitalSignatureCollection;
import com.aspose.words.DigitalSignatureUtil;
import org.testng.annotations.Test;

import java.util.Iterator;

public class ExDigitalSignatureCollection extends ApiExampleBase {
    @Test
    public void iterator() throws Exception {
        //ExStart
        //ExFor:DigitalSignatureCollection.GetEnumerator
        //ExSummary:Shows how to load and enumerate all digital signatures of a document.
        DigitalSignatureCollection digitalSignatures =
                DigitalSignatureUtil.loadSignatures(getMyDir() + "Digitally signed.docx");

        Iterator<DigitalSignature> enumerator = digitalSignatures.iterator();
        while (enumerator.hasNext()) {
            // Do something useful
            DigitalSignature ds = enumerator.next();

            if (ds != null)
                System.out.println(ds.toString());
        }
        //ExEnd
    }
}
