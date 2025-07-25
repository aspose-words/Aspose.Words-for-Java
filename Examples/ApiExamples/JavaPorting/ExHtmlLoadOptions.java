// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.HtmlLoadOptions;
import com.aspose.words.Document;
import org.testng.Assert;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.ImageType;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.WarningSource;
import com.aspose.words.WarningType;
import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningInfo;
import java.util.ArrayList;
import com.aspose.words.HtmlFixedSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.CertificateHolder;
import com.aspose.words.SignOptions;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.words.DigitalSignatureUtil;
import com.aspose.words.LoadFormat;
import com.aspose.words.HtmlControlType;
import com.aspose.words.NodeCollection;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.FormField;
import com.aspose.words.BlockImportMode;
import org.testng.annotations.DataProvider;


@Test
class ExHtmlLoadOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test (dataProvider = "supportVmlDataProvider")
    public void supportVml(boolean supportVml) throws Exception
    {
        //ExStart
        //ExFor:HtmlLoadOptions
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
            Assert.assertEquals(ImageType.JPEG, ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getImageData().getImageType());
        else
            Assert.assertEquals(ImageType.PNG, ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getImageData().getImageType());
        //ExEnd

        Shape imageShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        if (supportVml)
            TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        else
            TestUtil.verifyImageInShape(400, 400, ImageType.PNG, imageShape);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "supportVmlDataProvider")
	public static Object[][] supportVmlDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    //ExStart
    //ExFor:HtmlLoadOptions.WebRequestTimeout
    //ExSummary:Shows how to set a time limit for web requests when loading a document with external resources linked by URLs.
    @Test //ExSkip
    public void webRequestTimeout() throws Exception
    {
        // Create a new HtmlLoadOptions object and verify its timeout threshold for a web request.
        HtmlLoadOptions options = new HtmlLoadOptions();

        // When loading an Html document with resources externally linked by a web address URL,
        // Aspose.Words will abort web requests that fail to fetch the resources within this time limit, in milliseconds.
        Assert.assertEquals(100000, options.getWebRequestTimeout());

        // Set a WarningCallback that will record all warnings that occur during loading.
        ListDocumentWarnings warningCallback = new ListDocumentWarnings();
        options.setWarningCallback(warningCallback);

        // Load such a document and verify that a shape with image data has been created.
        // This linked image will require a web request to load, which will have to complete within our time limit.
        String html = $"\n                <html>\n                    <img src=\"{ImageUrl}\" alt=\"Aspose logo\" style=\"width:400px;height:400px;\">\n                </html>\n            ";

        // Set an unreasonable timeout limit and try load the document again.
        options.setWebRequestTimeout(0);
        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(html)), options);
        Assert.assertEquals(2, warningCallback.warnings().size());

        // A web request that fails to obtain an image within the time limit will still produce an image.
        // However, the image will be the red 'x' that commonly signifies missing images.
        Shape imageShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertEquals(924, imageShape.getImageData().getImageBytes().length);

        // We can also configure a custom callback to pick up any warnings from timed out web requests.
        Assert.assertEquals(WarningSource.HTML, warningCallback.warnings().get(0).getSource());
        Assert.assertEquals(WarningType.DATA_LOSS, warningCallback.warnings().get(0).getWarningType());
        Assert.assertEquals("Couldn't load a resource from \'{ImageUrl}\'.", warningCallback.warnings().get(0).getDescription());

        Assert.assertEquals(WarningSource.HTML, warningCallback.warnings().get(1).getSource());
        Assert.assertEquals(WarningType.DATA_LOSS, warningCallback.warnings().get(1).getWarningType());
        Assert.assertEquals("Image has been replaced with a placeholder.", warningCallback.warnings().get(1).getDescription());

        doc.save(getArtifactsDir() + "HtmlLoadOptions.WebRequestTimeout.docx");
    }

    /// <summary>
    /// Stores all warnings that occur during a document loading operation in a List.
    /// </summary>
    private static class ListDocumentWarnings implements IWarningCallback
    {
        public void warning(WarningInfo info)
        {
            mWarnings.add(info);
        }

        public ArrayList<WarningInfo> warnings() { 
            return mWarnings;
        }

        private /*final*/ ArrayList<WarningInfo> mWarnings = new ArrayList<WarningInfo>();
    }
    //ExEnd

    @Test
    public void loadHtmlFixed() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions(); { saveOptions.setSaveFormat(SaveFormat.HTML_FIXED); }

        doc.save(getArtifactsDir() + "HtmlLoadOptions.Fixed.html", saveOptions);

        HtmlLoadOptions loadOptions = new HtmlLoadOptions();

        ListDocumentWarnings warningCallback = new ListDocumentWarnings();
        loadOptions.setWarningCallback(warningCallback);

        doc = new Document(getArtifactsDir() + "HtmlLoadOptions.Fixed.html", loadOptions);
        Assert.assertEquals(1, warningCallback.warnings().size());

        Assert.assertEquals(WarningSource.HTML, warningCallback.warnings().get(0).getSource());
        Assert.assertEquals(WarningType.MAJOR_FORMATTING_LOSS, warningCallback.warnings().get(0).getWarningType());
        Assert.assertEquals("The document is fixed-page HTML. Its structure may not be loaded correctly.", warningCallback.warnings().get(0).getDescription());
    }

    @Test
    public void encryptedHtml() throws Exception
    {
        //ExStart
        //ExFor:HtmlLoadOptions.#ctor(String)
        //ExSummary:Shows how to encrypt an Html document, and then open it using a password.
        // Create and sign an encrypted HTML document from an encrypted .docx.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setComments("Comment");
            signOptions.setSignTime(new Date);
            signOptions.setDecryptionPassword("docPassword");
        }

        String inputFileName = getMyDir() + "Encrypted.docx";
        String outputFileName = getArtifactsDir() + "HtmlLoadOptions.EncryptedHtml.html";
        DigitalSignatureUtil.sign(inputFileName, outputFileName, certificateHolder, signOptions);

        // To load and read this document, we will need to pass its decryption
        // password using a HtmlLoadOptions object.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions("docPassword");

        Assert.assertEquals(signOptions.getDecryptionPassword(), loadOptions.getPassword());

        Document doc = new Document(outputFileName, loadOptions);

        Assert.assertEquals("Test encrypted document.", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void baseUri() throws Exception
    {
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
        Shape imageShape = (Shape)doc.getChildNodes(NodeType.SHAPE, true).get(0);
        Assert.assertTrue(imageShape.isImage());

        // This output document will display the image that was missing.
        doc.save(getArtifactsDir() + "HtmlLoadOptions.BaseUri.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "HtmlLoadOptions.BaseUri.docx");

        Assert.assertTrue(((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getImageData().getImageBytes().length > 0);
    }

    @Test
    public void getSelectAsSdt() throws Exception
    {
        //ExStart
        //ExFor:HtmlLoadOptions.PreferredControlType
        //ExFor:HtmlControlType
        //ExSummary:Shows how to set preferred type of document nodes that will represent imported <input> and <select> elements.
        final String HTML = "\n                <html>\n                    <select name='ComboBox' size='1'>\n                        <option value='val1'>item1</option>\n                        <option value='val2'></option>\n                    </select>\n                </html>\n            ";

        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
        htmlLoadOptions.setPreferredControlType(HtmlControlType.STRUCTURED_DOCUMENT_TAG);

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(HTML)), htmlLoadOptions);
        NodeCollection nodes = doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true);

        StructuredDocumentTag tag = (StructuredDocumentTag) nodes.get(0);
        //ExEnd

        Assert.assertEquals(2, tag.getListItems().getCount());

        Assert.assertEquals("val1", tag.getListItems().get(0).getValue());
        Assert.assertEquals("val2", tag.getListItems().get(1).getValue());
    }

    @Test
    public void getInputAsFormField() throws Exception
    {
        final String HTML = "\n                <html>\n                    <input type='text' value='Input value text' />\n                </html>\n            ";

        // By default, "HtmlLoadOptions.PreferredControlType" value is "HtmlControlType.FormField".
        // So, we do not set this value.
        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(HTML)), htmlLoadOptions);
        NodeCollection nodes = doc.getChildNodes(NodeType.FORM_FIELD, true);

        Assert.assertEquals(1, nodes.getCount());

        FormField formField = (FormField) nodes.get(0);
        Assert.assertEquals("Input value text", formField.getResult());
    }

    @Test (dataProvider = "ignoreNoscriptElementsDataProvider")
    public void ignoreNoscriptElements(boolean ignoreNoscriptElements) throws Exception
    {
        //ExStart
        //ExFor:HtmlLoadOptions.IgnoreNoscriptElements
        //ExSummary:Shows how to ignore <noscript> HTML elements.
        final String HTML = "\n                <html>\n                  <head>\n                    <title>NOSCRIPT</title>\n                      <meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\n                      <script type=\"text/javascript\">\n                        alert(\"Hello, world!\");\n                      </script>\n                  </head>\n                <body>\n                  <noscript><p>Your browser does not support JavaScript!</p></noscript>\n                </body>\n                </html>";

        HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
        htmlLoadOptions.setIgnoreNoscriptElements(ignoreNoscriptElements);

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(HTML)), htmlLoadOptions);
        doc.save(getArtifactsDir() + "HtmlLoadOptions.IgnoreNoscriptElements.pdf");
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "ignoreNoscriptElementsDataProvider")
	public static Object[][] ignoreNoscriptElementsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "usePdfDocumentForIgnoreNoscriptElementsDataProvider")
    public void usePdfDocumentForIgnoreNoscriptElements(boolean ignoreNoscriptElements) throws Exception
    {
        ignoreNoscriptElements(ignoreNoscriptElements);

        Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "HtmlLoadOptions.IgnoreNoscriptElements.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);

        Assert.That(textAbsorber.Text, assertEquals(ignoreNoscriptElements ? "" : "Your browser does not support JavaScript!", );
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "usePdfDocumentForIgnoreNoscriptElementsDataProvider")
	public static Object[][] usePdfDocumentForIgnoreNoscriptElementsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "blockImportDataProvider")
    public void blockImport(/*BlockImportMode*/int blockImportMode) throws Exception
    {
        //ExStart
        //ExFor:HtmlLoadOptions.BlockImportMode
        //ExFor:BlockImportMode
        //ExSummary:Shows how properties of block-level elements are imported from HTML-based documents.
        final String HTML = "\n            <html>\n                <div style='border:dotted'>\n                    <div style='border:solid'>\n                        <p>paragraph 1</p>\n                        <p>paragraph 2</p>\n                    </div>\n                </div>\n            </html>";
        MemoryStream stream = new MemoryStream(Encoding.getUTF8().getBytes(HTML));

        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        // Set the new mode of import HTML block-level elements.
        loadOptions.setBlockImportMode(blockImportMode);

        Document doc = new Document(stream, loadOptions);
        doc.save(getArtifactsDir() + "HtmlLoadOptions.BlockImport.docx");
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "blockImportDataProvider")
	public static Object[][] blockImportDataProvider() throws Exception
	{
		return new Object[][]
		{
			{BlockImportMode.PRESERVE},
			{BlockImportMode.MERGE},
		};
	}

    @Test
    public void fontFaceRules() throws Exception
    {
        //ExStart:FontFaceRules
        //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
        //ExFor:HtmlLoadOptions.SupportFontFaceRules
        //ExSummary:Shows how to load declared "@font-face" rules.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.setSupportFontFaceRules(true);
        Document doc = new Document(getMyDir() + "Html with FontFace.html", loadOptions);

        Assert.assertEquals("Squarish Sans CT Regular", doc.getFontInfos().get(0).getName());
        //ExEnd:FontFaceRules
    }
}
