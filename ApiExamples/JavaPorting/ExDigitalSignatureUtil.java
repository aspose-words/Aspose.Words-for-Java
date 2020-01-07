// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DigitalSignatureUtil;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.words.DigitalSignatureCollection;
import org.testng.Assert;
import com.aspose.words.CertificateHolder;
import com.aspose.words.SignOptions;
import com.aspose.ms.System.DateTime;
import com.aspose.words.LoadOptions;
import com.aspose.words.IncorrectPasswordException;
import com.aspose.ms.NUnit.Framework.msAssert;


@Test
public class ExDigitalSignatureUtil extends ApiExampleBase
{
    @Test
    public void removeAllSignatures() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignatureUtil
        //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
        //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
        //ExSummary:Shows how to remove every signature from a document.
        // Remove all signatures from the document using string parameters
        Document doc = new Document(getMyDir() + "Document.DigitalSignature.docx");
        String outFileName = getArtifactsDir() + "Document.NoSignatures.FromString.docx";

        DigitalSignatureUtil.removeAllSignatures(doc.getOriginalFileName(), outFileName);

        // Remove all signatures from the document using stream parameters
        Stream streamIn = new FileStream(getMyDir() + "Document.DigitalSignature.docx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            Stream streamOut = new FileStream(getArtifactsDir() + "Document.NoSignatures.FromStream.docx", FileMode.CREATE);
            try /*JAVA: was using*/
            {
                DigitalSignatureUtil.removeAllSignaturesInternal(streamIn, streamOut);
            }
            finally { if (streamOut != null) streamOut.close(); } 
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd
    }

    @Test
    public void loadSignatures() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
        //ExFor:DigitalSignatureUtil.LoadSignatures(String)
        //ExSummary:Shows how to load all existing signatures from a document.
        // Load all signatures from the document using string parameters
        DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.loadSignatures(getMyDir() + "Document.DigitalSignature.docx");
        Assert.assertNotNull(digitalSignatures);
        
        // Load all signatures from the document using stream parameters
        Stream stream = new FileStream(getMyDir() + "Document.DigitalSignature.docx", FileMode.OPEN);
        digitalSignatures = DigitalSignatureUtil.loadSignaturesInternal(stream);
        //ExEnd

        stream.close();
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

        Stream streamIn = new FileStream(getMyDir() + "Document.DigitalSignature.docx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            Stream streamOut = new FileStream(getArtifactsDir() + "Document.DigitalSignature.docx", FileMode.OPEN_OR_CREATE);
            try /*JAVA: was using*/
            {
                DigitalSignatureUtil.signInternal(streamIn, streamOut, certificateHolder, signOptions);
            }
            finally { if (streamOut != null) streamOut.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd
    }

    @Test (description = "WORDSNET-13036, WORDSNET-16868")
    public void signDocumentObfuscationBug() throws Exception
    {
        CertificateHolder ch = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        Document doc = new Document(getMyDir() + "TestRepeatingSection.docx");
        String outputFileName = getArtifactsDir() + "TestRepeatingSection.Signed.doc";

        SignOptions signOptions = new SignOptions(); { signOptions.setComments("Comment"); signOptions.setSignTime(DateTime.getNow()); }

        DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, ch, signOptions);
    }

    @Test (description = "WORDSNET-16868")
    public void incorrectPasswordForDecrypting() throws Exception
    {
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        Document doc = new Document(getMyDir() + "Document.Encrypted.docx", new LoadOptions("docPassword"));
        String outputFileName = getArtifactsDir() + "Document.Encrypted.docx";

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
    public void signDocumentWithDecryptionPassword() throws Exception
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
        String inputFileName = getMyDir() + "Document.Encrypted.docx";
        String outputFileName = getArtifactsDir() + "Document.Encrypted.docx";

        DigitalSignatureUtil.sign(inputFileName, outputFileName, certificateHolder, signOptions);
        //ExEnd

        // Open encrypted document from a file
        LoadOptions loadOptions = new LoadOptions("docPassword");
        msAssert.areEqual(signOptions.getDecryptionPassword(),loadOptions.getPassword());

        Document signedDoc = new Document(outputFileName, loadOptions);

        // Check that encrypted document was successfully signed
        DigitalSignatureCollection signatures = signedDoc.getDigitalSignatures();
        if (signatures.isValid() && (signatures.getCount() > 0))
        {
            //The document was signed successfully
            Assert.Pass();
        }
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
        Document doc = new Document(getMyDir() + "Document.DigitalSignature.docx");
        String outputFileName = getArtifactsDir() + "Document.DigitalSignature.docx";

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
