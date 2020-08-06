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
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Date;

@Test
class ExHtmlLoadOptions extends ApiExampleBase {
    @Test(dataProvider = "supportVmlDataProvider")
    public void supportVml(boolean doSupportVml) throws Exception {
        //ExStart
        //ExFor:HtmlLoadOptions.#ctor
        //ExFor:HtmlLoadOptions.SupportVml
        //ExSummary:Shows how to support VML while parsing a document.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        // If value is true, then we take VML code into account while parsing the loaded document
        loadOptions.setSupportVml(doSupportVml);

        // This document contains an image within "<!--[if gte vml 1]>" tags, and another different image within "<![if !vml]>" tags
        // Upon loading the document, only the contents of the first tag will be shown if VML is enabled,
        // and only the contents of the second tag will be shown otherwise
        Document doc = new Document(getMyDir() + "VML conditional.htm", loadOptions);

        // Only one of the two unique images will be loaded, depending on the value of loadOptions.SupportVml
        Shape imageShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        if (doSupportVml)
            TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        else
            TestUtil.verifyImageInShape(400, 400, ImageType.PNG, imageShape);
        //ExEnd
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

    //ExStart
    //ExFor:HtmlLoadOptions.WebRequestTimeout
    //ExSummary:Shows how to set a time limit for web requests that will occur when loading an html document which links external resources.
    @Test //ExSkip
    public void webRequestTimeout() throws Exception {
        // Create a new HtmlLoadOptions object and verify its timeout threshold for a web request
        HtmlLoadOptions options = new HtmlLoadOptions();

        // When loading an Html document with resources externally linked by a web address URL,
        // web requests that fetch these resources that fail to complete within this time limit will be aborted
        Assert.assertEquals(100000, options.getWebRequestTimeout());

        // Set a WarningCallback that will record all warnings that occur during loading
        ListDocumentWarnings warningCallback = new ListDocumentWarnings();
        options.setWarningCallback(warningCallback);

        // Load such a document and verify that a shape with image data has been created,
        // provided the request to get that image took place within the timeout limit
        String html = "\r\n<html>\r\n<img src=\"{AsposeLogoUrl}\" alt=\"Aspose logo\" style=\"width:400px;height:400px;\">\r\n</html>\r\n";

        Document doc = new Document(new FileInputStream(html), options);
        Shape imageShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(7498, imageShape.getImageData().getImageBytes().length);
        Assert.assertEquals(0, warningCallback.warnings().size());

        // Set an unreasonable timeout limit and load the document again
        options.setWebRequestTimeout(0);
        doc = new Document(new FileInputStream(html), options);

        // If a request fails to complete within the timeout limit, a shape with image data will still be produced
        // However, the image will be the red 'x' that commonly signifies missing images
        imageShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertEquals(924, imageShape.getImageData().getImageBytes().length);

        // A timeout like this will also accumulate warnings that can be picked up by a WarningCallback implementation
        Assert.assertEquals(WarningSource.HTML, warningCallback.warnings().get(0).getSource());
        Assert.assertEquals(WarningType.DATA_LOSS, warningCallback.warnings().get(0).getWarningType());
        Assert.assertEquals("Couldn't load a resource from \'{AsposeLogoUrl}\'.", warningCallback.warnings().get(0).getDescription());

        Assert.assertEquals(WarningSource.HTML, warningCallback.warnings().get(1).getSource());
        Assert.assertEquals(WarningType.DATA_LOSS, warningCallback.warnings().get(1).getWarningType());
        Assert.assertEquals("Image has been replaced with a placeholder.", warningCallback.warnings().get(1).getDescription());

        doc.save(getArtifactsDir() + "HtmlLoadOptions.WebRequestTimeout.docx");
    }

    /// <summary>
    /// Stores all warnings occuring during a document loading operation in a list.
    /// </summary>
    private static class ListDocumentWarnings implements IWarningCallback {
        public void warning(WarningInfo info) {
            mWarnings.add(info);
        }

        public ArrayList<WarningInfo> warnings() {
            return mWarnings;
        }

        private ArrayList<WarningInfo> mWarnings = new ArrayList<>();
    }
    //ExEnd

    @Test
    public void encryptedHtml() throws Exception {
        //ExStart
        //ExFor:HtmlLoadOptions.#ctor(String)
        //ExSummary:Shows how to encrypt an Html document and then open it using a password.
        // Create and sign an encrypted html document from an encrypted .docx
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

        // This .html document will need a password to be decrypted, opened and have its contents accessed
        // The password is specified by HtmlLoadOptions.Password
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
        //ExSummary:Shows how to specify a base URI when opening an html document.
        // If we want to load an .html document which contains an image linked by a relative URI
        // while the image is in a different location, we will need to resolve the relative URI into an absolute one
        // by creating an HtmlLoadOptions and providing a base URI
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.HTML, "", getImageDir());
        Document doc = new Document(getMyDir() + "Missing image.html", loadOptions);

        // While the image was broken in the input .html, it was successfully found in our base URI
        Shape imageShape = (Shape) doc.getChildNodes(NodeType.SHAPE, true).get(0);
        Assert.assertTrue(imageShape.isImage());

        // The image will be displayed correctly by the output document
        doc.save(getArtifactsDir() + "HtmlLoadOptions.BaseUri.docx");
        //ExEnd
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

        Document doc = new Document(new ByteArrayInputStream(html.getBytes("UTF-8")), htmlLoadOptions);
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

        // By default "HtmlLoadOptions.PreferredControlType" value is "HtmlControlType.FormField"
        // So, we do not set this value
        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();

        Document doc = new Document(new ByteArrayInputStream(html.getBytes("UTF-8")), htmlLoadOptions);
        NodeCollection nodes = doc.getChildNodes(NodeType.FORM_FIELD, true);

        Assert.assertEquals(nodes.getCount(), 1);

        FormField formField = (FormField) nodes.get(0);
        Assert.assertEquals(formField.getResult(), "Input value text");
    }
}
