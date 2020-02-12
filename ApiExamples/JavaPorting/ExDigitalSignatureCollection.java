// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.DigitalSignatureCollection;
import com.aspose.words.DigitalSignatureUtil;
import java.util.Iterator;
import com.aspose.words.DigitalSignature;
import com.aspose.ms.System.msConsole;


@Test
public class ExDigitalSignatureCollection extends ApiExampleBase
{
    @Test
    public void iterator()
    {
        //ExStart
        //ExFor:DigitalSignatureCollection.GetEnumerator
        //ExSummary:Shows how to load and enumerate all digital signatures of a document.
        DigitalSignatureCollection digitalSignatures =
            DigitalSignatureUtil.loadSignatures(getMyDir() + "Digitally signed.docx");

        Iterator<DigitalSignature> enumerator = digitalSignatures.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                // Do something useful
                DigitalSignature ds = enumerator.next();

                if (ds != null)
                    System.out.println(ds.toString());
            }
        }
        finally { if (enumerator != null) enumerator.close(); }
        //ExEnd
    }
}
