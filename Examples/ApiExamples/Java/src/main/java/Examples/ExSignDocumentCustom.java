package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import TestData.TestClasses.SignPersonTestClass;
import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.UUID;

/// <summary>
/// This example demonstrates how to add new signature line to the document and sign it with your personal signature <see cref="SignDocument"/>.
/// </summary>
@Test
public class ExSignDocumentCustom extends ApiExampleBase {
    //ExStart
    //ExFor:CertificateHolder
    //ExFor:SignatureLineOptions.Signer
    //ExFor:SignatureLineOptions.SignerTitle
    //ExFor:SignatureLine.Id
    //ExFor:SignOptions.SignatureLineId
    //ExFor:SignOptions.SignatureLineImage
    //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
    //ExSummary:Shows how to add a signature line to a document, and then sign it using a digital certificate.
    @Test(description = "WORDSNET-16868") //ExSkip
    public static void sign() throws Exception {
        String signPersonName = "Ron Williams";
        String srcDocumentPath = getMyDir() + "Document.docx";
        String dstDocumentPath = getArtifactsDir() + "SignDocumentCustom.Sign.docx";
        String certificatePath = getMyDir() + "morzal.pfx";
        String certificatePassword = "aw";

        // We need to create simple list with test signers for this example.
        createSignPersonData();
        System.out.println("Test data successfully added!");

        // Get sign person object by name of the person who must sign a document.
        // This an example, in real use case you would return an object from a database.
        SignPersonTestClass signPersonInfo = gSignPersonList.stream().filter(x -> x.getName() == signPersonName).findFirst().get();

        if (signPersonInfo != null) {
            signDocument(srcDocumentPath, dstDocumentPath, signPersonInfo, certificatePath, certificatePassword);
            System.out.println("Document successfully signed!");
        } else {
            System.out.println("Sign person does not exist, please check your parameters.");
            Assert.fail(); //ExSkip
        }

        // Now do something with a signed document, for example, save it to your database.
        // Use 'new Document(dstDocumentPath)' for loading a signed document.
    }

    /// <summary>
    /// Signs the document obtained at the source location and saves it to the specified destination.
    /// </summary>
    private static void signDocument(final String srcDocumentPath, final String dstDocumentPath,
                                     final SignPersonTestClass signPersonInfo, final String certificatePath,
                                     final String certificatePassword) throws Exception {
        // Create new document instance based on a test file that we need to sign.
        Document document = new Document(srcDocumentPath);
        DocumentBuilder builder = new DocumentBuilder(document);

        // Add info about responsible person who sign a document.
        SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
        signatureLineOptions.setSigner(signPersonInfo.getName());
        signatureLineOptions.setSignerTitle(signPersonInfo.getPosition());

        // Add signature line for responsible person who sign a document.
        SignatureLine signatureLine = builder.insertSignatureLine(signatureLineOptions).getSignatureLine();
        signatureLine.setId(signPersonInfo.getPersonId());

        // Save a document with line signatures into temporary file for future signing.
        builder.getDocument().save(dstDocumentPath);

        // Create holder of certificate instance based on your personal certificate.
        // This is the test certificate generated for this example.
        CertificateHolder certificateHolder = CertificateHolder.create(certificatePath, certificatePassword);

        // Link our signature line with personal signature.
        SignOptions signOptions = new SignOptions();
        signOptions.setSignatureLineId(signPersonInfo.getPersonId());
        signOptions.setSignatureLineImage(signPersonInfo.getImage());

        // Sign a document which contains signature line with personal certificate.
        DigitalSignatureUtil.sign(dstDocumentPath, dstDocumentPath, certificateHolder, signOptions);
    }

    /// <summary>
    /// Create test data that contains info about sing persons.
    /// </summary>
    private static void createSignPersonData() throws IOException {
        InputStream inputStream = new FileInputStream(getImageDir() + "Logo.jpg");

        gSignPersonList = new ArrayList<>();
        gSignPersonList.add(new SignPersonTestClass(UUID.randomUUID(), "Ron Williams", "Chief Executive Officer",
                DocumentHelper.getBytesFromStream(inputStream)));
        gSignPersonList.add(new SignPersonTestClass(UUID.randomUUID(), "Stephen Morse", "Head of Compliance",
                DocumentHelper.getBytesFromStream(inputStream)));
    }

    private static ArrayList<SignPersonTestClass> gSignPersonList;
    //ExEnd
}
