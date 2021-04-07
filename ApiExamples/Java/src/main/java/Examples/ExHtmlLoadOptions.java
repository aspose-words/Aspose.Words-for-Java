package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.pdf.TextAbsorber;
import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Date;

public class ExHtmlLoadOptions extends ApiExampleBase {
    @Test(dataProvider = "supportVmlDataProvider")
    public void supportVml(boolean supportVml) throws Exception {
        //ExStart
        //ExFor:HtmlLoadOptions.#ctor
        //ExFor:HtmlLoadOptions.SupportVml
        //ExSummary:Shows how to support conditional comments while loading an HTML document.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        // If the value is true, then we take VML code into account while parsing the loaded document.
        loadOptions.setSupportVml(supportVml);

        // This document contains a JPEG image within "<!--[if gte vml 1]>" tags,
        // and a different PNG image within "<![if !vml]>" tags.
        // If we set the "SupportVml" flag to "true", then Aspose.Words will load the JPEG.
        // If we set this flag to "false", then Aspose.Words will only load the PNG.
        Document doc = new Document(getMyDir() + "VML conditional.htm", loadOptions);

        if (supportVml)
            Assert.assertEquals(ImageType.JPEG, ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getImageData().getImageType());
        else
            Assert.assertEquals(ImageType.PNG, ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getImageData().getImageType());
        //ExEnd

        Shape imageShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        if (supportVml)
            TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        else
            TestUtil.verifyImageInShape(400, 400, ImageType.PNG, imageShape);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "supportVmlDataProvider")
    public static Object[][] supportVmlDataProvider() throws Exception {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    /// <summary>
    /// Stores all warnings that occur during a document loading operation in a List.
    /// </summary>
    private static class ListDocumentWarnings implements IWarningCallback {
        public void warning(WarningInfo info) {
            mWarnings.add(info);
        }

        public ArrayList<WarningInfo> warnings() {
            return mWarnings;
        }

        private final ArrayList<WarningInfo> mWarnings = new ArrayList<>();
    }

    @Test
    public void encryptedHtml() throws Exception {
        //ExStart
        //ExFor:HtmlLoadOptions.#ctor(String)
        //ExSummary:Shows how to encrypt an Html document, and then open it using a password.
        // Create and sign an encrypted HTML document from an encrypted .docx.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setComments("Comment");
            signOptions.setSignTime(new Date());
            signOptions.setDecryptionPassword("docPassword");
        }

        String inputFileName = getMyDir() + "Encrypted.docx";
        String outputFileName = getArtifactsDir() + "HtmlLoadOptions.EncryptedHtml.html";
        DigitalSignatureUtil.sign(inputFileName, outputFileName, certificateHolder, signOptions);

        // To load and read this document, we will need to pass its decryption
        // password using a HtmlLoadOptions object.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions("docPassword");
        Assert.assertEquals(loadOptions.getPassword(), signOptions.getDecryptionPassword());

        Document doc = new Document(outputFileName, loadOptions);
        Assert.assertEquals(doc.getText().trim(), "Test encrypted document.");
        //ExEnd
    }

    @Test
    public void baseUri() throws Exception {
        //ExStart
        //ExFor:HtmlLoadOptions.#ctor(LoadFormat,String,String)
        //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
        //ExFor:LoadOptions.LoadFormat
        //ExFor:LoadFormat
        //ExSummary:Shows how to specify a base URI when opening an html document.
        // Suppose we want to load an .html document that contains an image linked by a relative URI
        // while the image is in a different location. In that case, we will need to resolve the relative URI into an absolute one.
        // We can provide a base URI using an HtmlLoadOptions object. 
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.HTML, "", getImageDir());

        Assert.assertEquals(LoadFormat.HTML, loadOptions.getLoadFormat());

        Document doc = new Document(getMyDir() + "Missing image.html", loadOptions);

        // While the image was broken in the input .html, our custom base URI helped us repair the link.
        Shape imageShape = (Shape) doc.getChildNodes(NodeType.SHAPE, true).get(0);
        Assert.assertTrue(imageShape.isImage());

        // This output document will display the image that was missing.
        doc.save(getArtifactsDir() + "HtmlLoadOptions.BaseUri.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "HtmlLoadOptions.BaseUri.docx");

        Assert.assertTrue(((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getImageData().getImageBytes().length > 0);
    }

    @Test
    public void getSelectAsSdt() throws Exception {
        //ExStart
        //ExFor:HtmlLoadOptions.PreferredControlType
        //ExSummary:Shows how to set preferred type of document nodes that will represent imported <input> and <select> elements.
        final String html = "\r\n<html>\r\n<select name='ComboBox' size='1'>\r\n"
                + "<option value='val1'>item1</option>\r\n<option value='val2'></option>\r\n</select>\r\n</html>\r\n";

        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
        htmlLoadOptions.setPreferredControlType(HtmlControlType.STRUCTURED_DOCUMENT_TAG);

        Document doc = new Document(new ByteArrayInputStream(html.getBytes(StandardCharsets.UTF_8)), htmlLoadOptions);
        NodeCollection nodes = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        StructuredDocumentTag tag = (StructuredDocumentTag) nodes.get(0);
        //ExEnd

        Assert.assertEquals(tag.getListItems().getCount(), 2);

        Assert.assertEquals(tag.getListItems().get(0).getValue(), "val1");
        Assert.assertEquals(tag.getListItems().get(1).getValue(), "val2");
    }

    @Test
    public void getInputAsFormField() throws Exception {
        final String html = "\r\n<html>\r\n<input type='text' value='Input value text' />\r\n</html>\r\n";

        // By default, "HtmlLoadOptions.PreferredControlType" value is "HtmlControlType.FormField".
        // So, we do not set this value.
        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();

        Document doc = new Document(new ByteArrayInputStream(html.getBytes(StandardCharsets.UTF_8)), htmlLoadOptions);
        NodeCollection nodes = doc.getChildNodes(NodeType.FORM_FIELD, true);

        Assert.assertEquals(nodes.getCount(), 1);

        FormField formField = (FormField) nodes.get(0);
        Assert.assertEquals(formField.getResult(), "Input value text");
    }

    @Test (dataProvider = "ignoreNoscriptElementsDataProvider")
    public void ignoreNoscriptElements(boolean ignoreNoscriptElements) throws Exception
    {
        //ExStart
        //ExFor:HtmlLoadOptions.IgnoreNoscriptElements
        //ExSummary:Shows how to ignore <noscript> HTML elements.
        final String html = "\r\n<html>\r\n<head>\r\n<title>NOSCRIPT</title>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\n<script type=\"text/javascript\">\r\nalert(\"Hello, world!\");\r\n</script>\r\n</head>\r\n<body>\r\n<noscript><p>Your browser does not support JavaScript!</p></noscript>\r\n</body>\r\n</html>";

        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
        htmlLoadOptions.setIgnoreNoscriptElements(ignoreNoscriptElements);

        Document doc = new Document(new ByteArrayInputStream(html.getBytes(StandardCharsets.UTF_8)), htmlLoadOptions);
        doc.save(getArtifactsDir() + "HtmlLoadOptions.IgnoreNoscriptElements.pdf");
        //ExEnd

        com.aspose.pdf.Document pdfDoc = new com.aspose.pdf.Document(getArtifactsDir() + "HtmlLoadOptions.IgnoreNoscriptElements.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.visit(pdfDoc);

        Assert.assertEquals(ignoreNoscriptElements ? "" : "Your browser does not support JavaScript!", textAbsorber.getText());

        pdfDoc.close();
    }

	@DataProvider(name = "ignoreNoscriptElementsDataProvider")
	public static Object[][] ignoreNoscriptElementsDataProvider() {
		return new Object[][]
		{
			{true},
			{false},
		};
	}
}
