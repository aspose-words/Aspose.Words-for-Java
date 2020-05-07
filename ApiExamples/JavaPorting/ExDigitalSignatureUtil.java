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
import org.testng.Assert;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.words.CertificateHolder;
import com.aspose.words.SignOptions;
import com.aspose.ms.System.DateTime;
import com.aspose.words.DigitalSignature;
import com.aspose.words.DigitalSignatureType;
import com.aspose.words.Document;
import com.aspose.words.LoadOptions;
import com.aspose.words.IncorrectPasswordException;


@Test
public class ExDigitalSignatureUtil extends ApiExampleBase
{
    @Test
    public void loadAndRemove() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignatureUtil
        //ExFor:DigitalSignatureUtil.LoadSignatures(String)
        //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
        //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
        //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
        //ExSummary:Shows how to load and remove digital signatures from a digitally signed document.
        // Load digital signatures via filename string to verify that the document is signed
        DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.loadSignatures(getMyDir() + "Digitally signed.docx");
        Assert.assertEquals(1, digitalSignatures.getCount());

        // Re-save the document to an output filename with all digital signatures removed
        DigitalSignatureUtil.removeAllSignatures(getMyDir() + "Digitally signed.docx", getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromString.docx");

        // Remove all signatures from the document using stream parameters
        Stream streamIn = new FileStream(getMyDir() + "Digitally signed.docx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            Stream streamOut = new FileStream(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx", FileMode.CREATE);
            try /*JAVA: was using*/
            {
                DigitalSignatureUtil.removeAllSignaturesInternal(streamIn, streamOut);
            }
            finally { if (streamOut != null) streamOut.close(); } 
        }
        finally { if (streamIn != null) streamIn.close(); }

        // We can also load a document's digital signatures via stream, which we will do to verify that all signatures have been removed
        Stream stream = new FileStream(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            digitalSignatures = DigitalSignatureUtil.loadSignaturesInternal(stream);
            Assert.assertEquals(0, digitalSignatures.getCount());
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd
    }
    
    @Test (description = "WORDSNET-16868")
    public void signDocument() throws Exception
    {
        //ExStart
        //ExFor:CertificateHolder
        //ExFor:CertificateHolder.Create(String, String)
        //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
        //ExFor:SignOptions.Comments
        //ExFor:SignOptions.SignTime
        //ExSummary:Shows how to sign documents using certificate holder and sign options.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        SignOptions signOptions = new SignOptions(); { signOptions.setComments("My comment"); signOptions.setSignTime(DateTime.getNow()); }

        Stream streamIn = new FileStream(getMyDir() + "Digitally signed.docx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            Stream streamOut = new FileStream(getArtifactsDir() + "DigitalSignatureUtil.SignDocument.docx", FileMode.OPEN_OR_CREATE);
            try /*JAVA: was using*/
            {
                DigitalSignatureUtil.signInternal(streamIn, streamOut, certificateHolder, signOptions);
            }
            finally { if (streamOut != null) streamOut.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd

        Stream stream = new FileStream(getArtifactsDir() + "DigitalSignatureUtil.SignDocument.docx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.loadSignaturesInternal(stream);
            Assert.assertEquals(1, digitalSignatures.getCount());

            DigitalSignature signature = digitalSignatures.get(0);

            Assert.assertTrue(signature.isValid());
            Assert.assertEquals(DigitalSignatureType.XML_DSIG, signature.getSignatureType());
            Assert.assertEquals(signOptions.getSignTimeInternal().toString(), signature.getSignTimeInternal().toString());
            Assert.assertEquals("My comment", signature.getComments());
        }
        finally { if (stream != null) stream.close(); }
    }

    @Test (description = "WORDSNET-13036, WORDSNET-16868")
    public void signDocumentObfuscationBug() throws Exception
    {
        CertificateHolder ch = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        Document doc = new Document(getMyDir() + "Structured document tags.docx");
        String outputFileName = getArtifactsDir() + "DigitalSignatureUtil.SignDocumentObfuscationBug.doc";

        SignOptions signOptions = new SignOptions(); { signOptions.setComments("Comment"); signOptions.setSignTime(DateTime.getNow()); }

        DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, ch, signOptions);
    }

    @Test (description = "WORDSNET-16868")
    public void incorrectDecryptionPassword() throws Exception
    {
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        Document doc = new Document(getMyDir() + "Encrypted.docx", new LoadOptions("docPassword"));
        String outputFileName = getArtifactsDir() + "DigitalSignatureUtil.IncorrectDecryptionPassword.docx";

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setComments("Comment");
            signOptions.setSignTime(DateTime.getNow());
            signOptions.setDecryptionPassword("docPassword1");
        }

        // Digitally sign encrypted with "docPassword" document in the specified path
        Assert.That(
            new TestDelegate(() => DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, certificateHolder, signOptions)),
            Throws.<IncorrectPasswordException>TypeOf(), "The document password is incorrect.");
    }

    @Test (description = "WORDSNET-16868")
    public void decryptionPassword() throws Exception
    {
        //ExStart
        //ExFor:CertificateHolder
        //ExFor:SignOptions.DecryptionPassword
        //ExFor:LoadOptions.Password
        //ExSummary:Shows how to sign encrypted document file.
        // Create certificate holder from a file
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setComments("Comment");
            signOptions.setSignTime(DateTime.getNow());
            signOptions.setDecryptionPassword("docPassword");
        }

        // Digitally sign encrypted with "docPassword" document in the specified path
        String inputFileName = getMyDir() + "Encrypted.docx";
        String outputFileName = getArtifactsDir() + "DigitalSignatureUtil.DecryptionPassword.docx";

        DigitalSignatureUtil.sign(inputFileName, outputFileName, certificateHolder, signOptions);
        //ExEnd

        // Open encrypted document from a file
        LoadOptions loadOptions = new LoadOptions("docPassword");
        Assert.assertEquals(signOptions.getDecryptionPassword(),loadOptions.getPassword());

        // Check that encrypted document was successfully signed
        Document signedDoc = new Document(outputFileName, loadOptions);
        DigitalSignatureCollection signatures = signedDoc.getDigitalSignatures();

        Assert.assertEquals(1, signatures.getCount());
        Assert.assertTrue(signatures.isValid());
    }

    @Test
    public void noArgumentsForSing() throws Exception
    {
        SignOptions signOptions = new SignOptions();
        {
            signOptions.setComments("");
            signOptions.setSignTime(DateTime.getNow());
            signOptions.setDecryptionPassword("");
        }

        Assert.That(() => DigitalSignatureUtil.sign("", "", null, signOptions),
            Throws.<IllegalArgumentException>TypeOf());
    }

    @Test
    public void noCertificateForSign() throws Exception
    {
        Document doc = new Document(getMyDir() + "Digitally signed.docx");
        String outputFileName = getArtifactsDir() + "DigitalSignatureUtil.NoCertificateForSign.docx";

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setComments("Comment");
            signOptions.setSignTime(DateTime.getNow());
            signOptions.setDecryptionPassword("docPassword");
        }

        Assert.That(() => DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, null, signOptions),
            Throws.<NullPointerException>TypeOf());
    }
}
