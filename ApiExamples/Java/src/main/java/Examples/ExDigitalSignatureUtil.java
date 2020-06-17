package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
    public void loadAndRemove() throws Exception {
        //ExStart
        //ExFor:DigitalSignatureUtil
        //ExFor:DigitalSignatureUtil.LoadSignatures(String)
        //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
        //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
        //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
        //ExSummary:Shows how to load and remove digital signatures from a digitally signed document.
        // Load digital signatures via filename string to verify that the document is signed
        DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.loadSignatures(getMyDir() + "Digitally signed.docx");
        Assert.assertEquals(digitalSignatures.getCount(), 1);

        // Re-save the document to an output filename with all digital signatures removed
        DigitalSignatureUtil.removeAllSignatures(getMyDir() + "Digitally signed.docx", getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromString.docx");

        // Remove all signatures from the document using stream parameters
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

        // We can also load a document's digital signatures via stream, which we will do to verify that all signatures have been removed
        InputStream stream = new FileInputStream(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx");
        try /*JAVA: was using*/ {
            digitalSignatures = DigitalSignatureUtil.loadSignatures(stream);
            Assert.assertEquals(0, digitalSignatures.getCount());
        } finally {
            if (stream != null) stream.close();
        }
        //ExEnd
    }

    @Test(description = "WORDSNET-16868")
    public void signDocument() throws Exception {
        //ExStart
        //ExFor:CertificateHolder
        //ExFor:CertificateHolder.Create(String, String)
        //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
        //ExFor:SignOptions.Comments
        //ExFor:SignOptions.SignTime
        //ExSummary:Shows how to sign documents using certificate holder and sign options.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        // By string:
        Document doc = new Document(getMyDir() + "Digitally signed.docx");
        String outputFileName = getArtifactsDir() + "DigitalSignatureUtil.SignDocument.docx";

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Encrypted");
        signOptions.setSignTime(new Date());

        DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, certificateHolder, signOptions);
        //ExEnd
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
        signOptions.setComments("Comment");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("docPassword1");

        // Digitally sign encrypted with "docPassword" document in the specified path
        try {
            DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, certificateHolder, signOptions);
        } catch (Exception e) {
            Assert.assertTrue(e instanceof IncorrectPasswordException);
            Assert.assertEquals(e.getMessage(), "The document password is incorrect.");
        }
    }

    @Test(description = "WORDSNET-16868")
    public void decryptionPassword() throws Exception {
        //ExStart
        //ExFor:CertificateHolder
        //ExFor:SignOptions.DecryptionPassword
        //ExFor:LoadOptions.Password
        //ExSummary:Shows how to sign encrypted document file.
        // Create certificate holder from a file
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Comment");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("docPassword");

        // Digitally sign encrypted with "docPassword" document in the specified path
        String inputFileName = getMyDir() + "Encrypted.docx";
        String outputFileName = getArtifactsDir() + "DigitalSignatureUtil.DecryptionPassword.docx";

        DigitalSignatureUtil.sign(inputFileName, outputFileName, certificateHolder, signOptions);
        //ExEnd

        // Open encrypted document from a file
        LoadOptions loadOptions = new LoadOptions("docPassword");
        Assert.assertEquals(loadOptions.getPassword(), signOptions.getDecryptionPassword());

        // Check that encrypted document was successfully signed
        Document signedDoc = new Document(outputFileName, loadOptions);
        DigitalSignatureCollection signatures = signedDoc.getDigitalSignatures();

        Assert.assertEquals(1, signatures.getCount());
        Assert.assertTrue(signatures.isValid());
    }

    @Test
    public void singInputStreamDocumentWithPasswordDecrypring() throws Exception {
        //ExStart
        //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
        //ExSummary:Shows how to sign encrypted document opened from a stream.
        FileInputStream streamIn = new FileInputStream(getMyDir() + "Digitally signed.docx");
        FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "DigitalSignatureUtil.SignDocument.docx");

        // Create certificate holder from a file
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Encrypted");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("docPassword");

        // Digitally sign encrypted with "docPassword" document in the specified path
        DigitalSignatureUtil.sign(streamIn, streamOut, certificateHolder, signOptions);
        //ExEnd

        // Open encrypted document from a file
        InputStream streamOutIn = new FileInputStream(getArtifactsDir() + "DigitalSignatureUtil.SignDocument.docx");
        try {
            DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.loadSignatures(streamOutIn);
            Assert.assertEquals(1, digitalSignatures.getCount());

            DigitalSignature signature = digitalSignatures.get(0);

            Assert.assertTrue(signature.isValid());
            Assert.assertEquals(DigitalSignatureType.XML_DSIG, signature.getSignatureType());
            Assert.assertEquals("Encrypted", signature.getComments());
        } finally {
            if (streamOutIn != null) streamOutIn.close();
        }
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
        String outputFileName = getArtifactsDir() + "DigitalSignatureUtil.NoCertificateForSign.docx";

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Comment");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("docPassword");

        Assert.assertThrows(NullPointerException.class, () -> DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, null, signOptions));
    }
}
