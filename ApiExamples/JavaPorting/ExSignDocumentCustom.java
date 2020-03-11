// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.ms.System.msConsole;
import ApiExamples.TestData.TestClasses.SignPersonTestClass;
import org.testng.Assert;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SignatureLineOptions;
import com.aspose.words.SignatureLine;
import com.aspose.words.CertificateHolder;
import com.aspose.words.SignOptions;
import com.aspose.words.DigitalSignatureUtil;
import java.awt.image.BufferedImage;
import com.aspose.ms.System.IO.MemoryStream;
import java.util.ArrayList;
import com.aspose.ms.System.Guid;
import com.aspose.BitmapPal;


/// <summary>
/// This example demonstrates how to add new signature line to the document and sign it with your personal signature <see cref="SignDocument"/>.
/// </summary>
@Test
public class ExSignDocumentCustom extends ApiExampleBase
{
    //ExStart
    //ExFor:CertificateHolder
    //ExFor:SignatureLineOptions.Signer
    //ExFor:SignatureLineOptions.SignerTitle
    //ExFor:SignatureLine.Id
    //ExFor:SignOptions.SignatureLineId
    //ExFor:SignOptions.SignatureLineImage
    //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
    //ExSummary:Demonstrates how to add new signature line to the document and sign it with personal signature using SignatureLineId.
    @Test (description = "WORDSNET-16868") //ExSkip
    public static void sign() throws Exception
    {
        String signPersonName = "Ron Williams";
        String srcDocumentPath = getMyDir() + "Document.docx";
        String dstDocumentPath = getArtifactsDir() + "SignDocumentCustom.Sign.docx";
        String certificatePath = getMyDir() + "morzal.pfx";
        String certificatePassword = "aw";

        // We need to create simple list with test signers for this example
        createSignPersonData();
        System.out.println("Test data successfully added!");

        // Get sign person object by name of the person who must sign a document
        // This an example, in real use case you would return an object from a database
        SignPersonTestClass signPersonInfo =
            (from c : gSignPersonList where c.Name == signPersonName select c).FirstOrDefault();

        if (signPersonInfo != null)
        {
            signDocument(srcDocumentPath, dstDocumentPath, signPersonInfo, certificatePath, certificatePassword);
            System.out.println("Document successfully signed!");
        }
        else
        {
            System.out.println("Sign person does not exist, please check your parameters.");
            Assert.fail(); //ExSkip
        }

        // Now do something with a signed document, for example, save it to your database
        // Use 'new Document(dstDocumentPath)' for loading a signed document
    }

    /// <summary>
    /// Signs the document obtained at the source location and saves it to the specified destination.
    /// </summary>
    private static void signDocument(String srcDocumentPath, String dstDocumentPath,
        SignPersonTestClass signPersonInfo, String certificatePath, String certificatePassword) throws Exception
    {
        // Create new document instance based on a test file that we need to sign
        Document document = new Document(srcDocumentPath);
        DocumentBuilder builder = new DocumentBuilder(document);

        // Add info about responsible person who sign a document
        SignatureLineOptions signatureLineOptions =
            new SignatureLineOptions(); { signatureLineOptions.setSigner(signPersonInfo.getName()); signatureLineOptions.setSignerTitle(signPersonInfo.getPosition()); }

        // Add signature line for responsible person who sign a document
        SignatureLine signatureLine = builder.insertSignatureLine(signatureLineOptions).getSignatureLine();
        signatureLine.setIdInternal(signPersonInfo.getPersonId());

        // Save a document with line signatures into temporary file for future signing
        builder.getDocument().save(dstDocumentPath);

        // Create holder of certificate instance based on your personal certificate
        // This is the test certificate generated for this example
        CertificateHolder certificateHolder = CertificateHolder.create(certificatePath, certificatePassword);

        // Link our signature line with personal signature
        SignOptions signOptions = new SignOptions();
        {
            signOptions.setSignatureLineId(signPersonInfo.getPersonId());
            signOptions.setSignatureLineImage(signPersonInfo.getImage());
        }

        // Sign a document which contains signature line with personal certificate
        DigitalSignatureUtil.sign(dstDocumentPath, dstDocumentPath, certificateHolder, signOptions);
    }

        /// <summary>
    /// Converting image file to bytes array.
    /// </summary>
    private static byte[] imageToByteArray(BufferedImage imageIn) throws Exception
    {
        MemoryStream ms = new MemoryStream();
        try /*JAVA: was using*/
        {
            imageIn.Save(ms, ImageFormat.Png);
            return ms.toArray();
        }
        finally { if (ms != null) ms.close(); }
    }
    
    /// <summary>
    /// Create test data that contains info about sing persons
    /// </summary>
    private static void createSignPersonData() throws Exception
    {
        gSignPersonList = new ArrayList<SignPersonTestClass>();
        {
                        gSignPersonList.add(new SignPersonTestClass(Guid.newGuid(), "Ron Williams", "Chief Executive Officer",
                imageToByteArray(BitmapPal.loadNativeImage(getImageDir() + "Logo.jpg"))));
                                        
                        gSignPersonList.add(new SignPersonTestClass(Guid.newGuid(), "Stephen Morse", "Head of Compliance",
                imageToByteArray(BitmapPal.loadNativeImage(getImageDir() + "Logo.jpg"))));
                                    }
    }

    private static ArrayList<SignPersonTestClass> gSignPersonList;
    //ExEnd
}
