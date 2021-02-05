package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;

public class ExDigitalSignatureUtil extends ApiExampleBase {
    @Test
    public void load() throws Exception {
        //ExStart
        //ExFor:DigitalSignatureUtil
        //ExFor:DigitalSignatureUtil.LoadSignatures(String)
        //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
        //ExSummary:Shows how to load signatures from a digitally signed document.
        // There are two ways of loading a signed document's collection of digital signatures using the DigitalSignatureUtil class.
        // 1 -  Load from a document from a local file system filename:
        DigitalSignatureCollection digitalSignatures =
                DigitalSignatureUtil.loadSignatures(getMyDir() + "Digitally signed.docx");

        // If this collection is nonempty, then we can verify that the document is digitally signed.
        Assert.assertEquals(1, digitalSignatures.getCount());

        // 2 -  Load from a document from a FileStream:
        InputStream stream = new FileInputStream(getMyDir() + "Digitally signed.docx");
        try {
            digitalSignatures = DigitalSignatureUtil.loadSignatures(stream);
            Assert.assertEquals(1, digitalSignatures.getCount());
        } finally {
            if (stream != null) stream.close();
        }
        //ExEnd
    }

    @Test
    public void remove() throws Exception {
        //ExStart
        //ExFor:DigitalSignatureUtil
        //ExFor:DigitalSignatureUtil.LoadSignatures(String)
        //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
        //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
        //ExSummary:Shows how to remove digital signatures from a digitally signed document.
        // There are two ways of using the DigitalSignatureUtil class to remove digital signatures
        // from a signed document by saving an unsigned copy of it somewhere else in the local file system.
        // 1 - Determine the locations of both the signed document and the unsigned copy by filename strings:
        DigitalSignatureUtil.removeAllSignatures(getMyDir() + "Digitally signed.docx",
                getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromString.docx");

        // 2 - Determine the locations of both the signed document and the unsigned copy by file streams:
        InputStream streamIn = new FileInputStream(getMyDir() + "Digitally signed.docx");
        try {
            OutputStream streamOut = new FileOutputStream(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx");
            try {
                DigitalSignatureUtil.removeAllSignatures(streamIn, streamOut);
            } finally {
                if (streamOut != null) streamOut.close();
            }
        } finally {
            if (streamIn != null) streamIn.close();
        }

        // Verify that both our output documents have no digital signatures.
        Assert.assertEquals(DigitalSignatureUtil.loadSignatures(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromString.docx").getCount(), 0);
        Assert.assertEquals(DigitalSignatureUtil.loadSignatures(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx").getCount(), 0);
        //ExEnd
    }

    @Test(description = "WORDSNET-16868, WORDSJAVA-2406", enabled = false)
    public void signDocument() throws Exception {
        //ExStart
        //ExFor:CertificateHolder
        //ExFor:CertificateHolder.Create(String, String)
        //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
        //ExFor:SignOptions.Comments
        //ExFor:SignOptions.SignTime
        //ExSummary:Shows how to digitally sign documents.
        // Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        // Create a comment and date which will be applied with our new digital signature.
        SignOptions signOptions = new SignOptions();
        {
            signOptions.setComments("My comment");
            signOptions.setSignTime(new Date());
        }

        // Take an unsigned document from the local file system via a file stream,
        // then create a signed copy of it determined by the filename of the output file stream.
        InputStream streamIn = new FileInputStream(getMyDir() + "Document.docx");
        try {
            OutputStream streamOut = new FileOutputStream(getArtifactsDir() + "DigitalSignatureUtil.SignDocument.docx");
            try {
                DigitalSignatureUtil.sign(streamIn, streamOut, certificateHolder, signOptions);
            } finally {
                if (streamOut != null) streamOut.close();
            }
        } finally {
            if (streamIn != null) streamIn.close();
        }
        //ExEnd

        InputStream stream = new FileInputStream(getArtifactsDir() + "DigitalSignatureUtil.SignDocument.docx");
        try {
            DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.loadSignatures(stream);
            Assert.assertEquals(1, digitalSignatures.getCount());

            DigitalSignature signature = digitalSignatures.get(0);

            Assert.assertTrue(signature.isValid());
            Assert.assertEquals(DigitalSignatureType.XML_DSIG, signature.getSignatureType());
            Assert.assertEquals(signOptions.getSignTime().toString(), signature.getSignTime().toString());
            Assert.assertEquals("My comment", signature.getComments());
        } finally {
            if (stream != null) stream.close();
        }
    }

    @Test(description = "WORDSNET-16868")
    public void decryptionPassword() throws Exception {
        //ExStart
        //ExFor:CertificateHolder
        //ExFor:SignOptions.DecryptionPassword
        //ExFor:LoadOptions.Password
        //ExSummary:Shows how to sign encrypted document file.
        // Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        // Create a comment, date, and decryption password which will be applied with our new digital signature.
        SignOptions signOptions = new SignOptions();
        {
            signOptions.setComments("Comment");
            signOptions.setSignTime(new Date());
            signOptions.setDecryptionPassword("docPassword");
        }

        // Set a local system filename for the unsigned input document, and an output filename for its new digitally signed copy.
        String inputFileName = getMyDir() + "Encrypted.docx";
        String outputFileName = getArtifactsDir() + "DigitalSignatureUtil.DecryptionPassword.docx";

        DigitalSignatureUtil.sign(inputFileName, outputFileName, certificateHolder, signOptions);
        //ExEnd

        // Open encrypted document from a file.
        LoadOptions loadOptions = new LoadOptions("docPassword");
        Assert.assertEquals(signOptions.getDecryptionPassword(), loadOptions.getPassword());

        // Check that encrypted document was successfully signed.
        Document signedDoc = new Document(outputFileName, loadOptions);
        DigitalSignatureCollection signatures = signedDoc.getDigitalSignatures();

        Assert.assertEquals(1, signatures.getCount());
        Assert.assertTrue(signatures.isValid());
    }

    @Test(description = "WORDSNET-13036, WORDSNET-16868")
    public void signDocumentObfuscationBug() throws Exception {
        CertificateHolder ch = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        Document doc = new Document(getMyDir() + "Structured document tags.docx");
        String outputFileName = getArtifactsDir() + "DigitalSignatureUtil.SignDocumentObfuscationBug.doc";

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Comment");
        signOptions.setSignTime(new Date());

        DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, ch, signOptions);
    }

    @Test(description = "WORDSNET-16868")
    public void incorrectDecryptionPassword() throws Exception {
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        Document doc = new Document(getMyDir() + "Encrypted.docx", new LoadOptions("docPassword"));
        String outputFileName = getArtifactsDir() + "DigitalSignatureUtil.IncorrectDecryptionPassword.docx";

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setComments("Comment");
            signOptions.setSignTime(new Date());
            signOptions.setDecryptionPassword("docPassword1");
        }

        Assert.assertThrows(IncorrectPasswordException.class, () -> DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, certificateHolder, signOptions));
    }

    @Test
    public void noArgumentsForSing() {
        SignOptions signOptions = new SignOptions();

        signOptions.setComments("");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("");

        Assert.assertThrows(IllegalArgumentException.class, () -> DigitalSignatureUtil.sign("", "", null, signOptions));
    }

    @Test
    public void noCertificateForSign() throws Exception {
        Document doc = new Document(getMyDir() + "Digitally signed.docx");

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Comment");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("docPassword");

        Assert.assertThrows(NullPointerException.class, () -> DigitalSignatureUtil.sign(doc.getOriginalFileName(),
                getArtifactsDir() + "DigitalSignatureUtil.NoCertificateForSign.docx", null, signOptions));
    }
}
