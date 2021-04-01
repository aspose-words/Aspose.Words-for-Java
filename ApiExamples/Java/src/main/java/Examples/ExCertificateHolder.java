package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.CertificateHolder;
import org.testng.annotations.Test;

import java.io.FileInputStream;

@Test
public class ExCertificateHolder extends ApiExampleBase {
    @Test
    public void create() throws Exception {
        //ExStart
        //ExFor:CertificateHolder.Create(Byte[], SecureString)
        //ExFor:CertificateHolder.Create(Byte[], String)
        //ExFor:CertificateHolder.Create(String, String, String)
        //ExSummary:Shows how to create CertificateHolder objects.
        // Below are four ways of creating CertificateHolder objects.
        // 1 -  Load a PKCS #12 file into a byte array and apply its password:
        byte[] certBytes = DocumentHelper.getBytesFromStream(new FileInputStream(getMyDir() + "morzal.pfx"));
        CertificateHolder.create(certBytes, "aw");

        // 2 -  Use a valid alias:
        CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", "c20be521-11ea-4976-81ed-865fbbfc9f24");

        // 3 -  Pass "null" as the alias in order to use the first available alias that returns a private key:
        CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", null);
        //ExEnd
    }
}
