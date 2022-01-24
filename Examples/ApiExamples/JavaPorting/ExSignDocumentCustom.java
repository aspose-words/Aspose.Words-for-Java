// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.ms.System.msString;
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
import com.aspose.ms.System.Guid;
import java.util.ArrayList;
import javax.imageio.ImageIO;


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
    //ExSummary:Shows how to add a signature line to a document, and then sign it using a digital certificate.
    @Test (description = "WORDSNET-16868") //ExSkip
    public static void sign() throws Exception
    {
        String signeeName = "Ron Williams";
        String srcDocumentPath = getMyDir() + "Document.docx";
        String dstDocumentPath = getArtifactsDir() + "SignDocumentCustom.Sign.docx";
        String certificatePath = getMyDir() + "morzal.pfx";
        String certificatePassword = "aw";

        createSignees();

        Signee signeeInfo = mSignees.Find(c => msString.equals(c.getName(), signeeName));

        if (signeeInfo != null)
            signDocument(srcDocumentPath, dstDocumentPath, signeeInfo, certificatePath, certificatePassword);
        else
            Assert.fail("Signee does not exist.");
    }

    /// <summary>
    /// Creates a copy of a source document signed using provided signee information and X509 certificate.
    /// </summary>
    private static void signDocument(String srcDocumentPath, String dstDocumentPath,
        Signee signeeInfo, String certificatePath, String certificatePassword) throws Exception
    {
        Document document = new Document(srcDocumentPath);
        DocumentBuilder builder = new DocumentBuilder(document);

        // Configure and insert a signature line, an object in the document that will display a signature that we sign it with.
        SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
        {
            signatureLineOptions.setSigner(signeeInfo.getName()); 
            signatureLineOptions.setSignerTitle(signeeInfo.getPosition());
        }

        SignatureLine signatureLine = builder.insertSignatureLine(signatureLineOptions).getSignatureLine();
        signatureLine.setIdInternal(signeeInfo.getPersonId());

        // First, we will save an unsigned version of our document.
        builder.getDocument().save(dstDocumentPath);

        CertificateHolder certificateHolder = CertificateHolder.create(certificatePath, certificatePassword);
        
        SignOptions signOptions = new SignOptions();
        {
            signOptions.setSignatureLineId(signeeInfo.getPersonId());
            signOptions.setSignatureLineImage(signeeInfo.getImage());
        }

        // Overwrite the unsigned document we saved above with a version signed using the certificate.
        DigitalSignatureUtil.sign(dstDocumentPath, dstDocumentPath, certificateHolder, signOptions);
    }

    /// <summary>
    /// Converts an image to a byte array.
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

    public static class Signee
    {
        public Guid getPersonId() { return mPersonId; }; public void setPersonId(Guid value) { mPersonId = value; };

        private Guid mPersonId;
        public String getName() { return mName; }; public void setName(String value) { mName = value; };

        private String mName;
        public String getPosition() { return mPosition; }; public void setPosition(String value) { mPosition = value; };

        private String mPosition;
        public byte[] getImage() { return mImage; }; public void setImage(byte[] value) { mImage = value; };

        private byte[] mImage;

        public Signee(Guid guid, String name, String position, byte[] image)
        {
            setPersonId(guid);
            setName(name);
            setPosition(position);
            setImage(image);
        }
    }

    private static void createSignees() throws Exception
    {
        mSignees = new ArrayList<Signee>();
        {
                        mSignees.add(new Signee(Guid.newGuid(), "Ron Williams", "Chief Executive Officer",
                imageToByteArray(ImageIO.read(getImageDir() + "Logo.jpg"))));
                                        
                        mSignees.add(new Signee(Guid.newGuid(), "Stephen Morse", "Head of Compliance",
                imageToByteArray(ImageIO.read(getImageDir() + "Logo.jpg"))));
                                    }
    }
    
    private static ArrayList<Signee> mSignees;
    //ExEnd
}
