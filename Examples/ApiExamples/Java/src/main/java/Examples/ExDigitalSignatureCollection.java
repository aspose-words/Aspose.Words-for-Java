package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.DigitalSignature;
import com.aspose.words.DigitalSignatureCollection;
import com.aspose.words.DigitalSignatureType;
import com.aspose.words.DigitalSignatureUtil;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.util.Iterator;

public class ExDigitalSignatureCollection extends ApiExampleBase {
    @Test
    public void iterator() throws Exception {
        //ExStart
        //ExFor:DigitalSignatureCollection.GetEnumerator
        //ExSummary:Shows how to print all the digital signatures of a signed document.
        DigitalSignatureCollection digitalSignatures =
                DigitalSignatureUtil.loadSignatures(getMyDir() + "Digitally signed.docx");

        Iterator<DigitalSignature> enumerator = digitalSignatures.iterator();
        while (enumerator.hasNext()) {
            DigitalSignature ds = enumerator.next();

            if (ds != null)
                System.out.println(ds.toString());
        }
        //ExEnd

        Assert.assertEquals(1, digitalSignatures.getCount());

        DigitalSignature signature = digitalSignatures.get(0);

        Assert.assertTrue(signature.isValid());
        Assert.assertEquals(DigitalSignatureType.XML_DSIG, signature.getSignatureType());
        Assert.assertEquals("Test Sign", signature.getComments());

        Assert.assertEquals(signature.getIssuerName(), signature.getIssuerName());
        Assert.assertEquals(signature.getSubjectName(), signature.getSubjectName());

        Assert.assertEquals("CN=VeriSign Class 3 Code Signing 2009-2 CA, " +
                "OU=Terms of use at https://www.verisign.com/rpa (c)09, " +
                "OU=VeriSign Trust Network, " +
                "O=\"VeriSign, Inc.\", " +
                "C=US", signature.getIssuerName());

        Assert.assertEquals("CN=Aspose Pty Ltd, " +
                "OU=Digital ID Class 3 - Microsoft Software Validation v2, " +
                "O=Aspose Pty Ltd, " +
                "L=Lane Cove, " +
                "S=New South Wales, " +
                "C=AU", signature.getSubjectName());
    }
}
