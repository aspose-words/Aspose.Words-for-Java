// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import org.testng.Assert;
import com.aspose.words.DocumentBase;
import com.aspose.words.GlossaryDocument;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.Run;
import com.aspose.words.Section;
import com.aspose.words.SaveFormat;
import com.aspose.words.Style;
import com.aspose.words.StyleType;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.NodeType;
import com.aspose.words.IResourceLoadingCallback;
import com.aspose.words.ResourceLoadingAction;
import com.aspose.words.ResourceLoadingArgs;
import com.aspose.words.ResourceType;
import com.aspose.ms.System.IO.File;


@Test
public class ExDocumentBase extends ApiExampleBase
{
    @Test
    public void constructor() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase
        //ExSummary:Shows how to initialize the subclasses of DocumentBase.
        Document doc = new Document();

        Assert.assertEquals(DocumentBase.class, doc.getClass().getSuperclass());

        GlossaryDocument glossaryDoc = new GlossaryDocument();
        doc.setGlossaryDocument(glossaryDoc);

        Assert.assertEquals(DocumentBase.class, glossaryDoc.getClass().getSuperclass());
        //ExEnd
    }

    @Test
    public void setPageColor() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase.PageColor
        //ExSummary:Shows how to set the background color for all pages of a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        doc.setPageColor(msColor.getLightGray());

        doc.save(getArtifactsDir() + "DocumentBase.SetPageColor.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBase.SetPageColor.docx");

        Assert.assertEquals(msColor.getLightGray().getRGB(), doc.getPageColor().getRGB());
    }

    @Test
    public void importNode() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase.ImportNode(Node, Boolean)
        //ExSummary:Shows how to import a node from one document to another.
        Document srcDoc = new Document();
        Document dstDoc = new Document();

        srcDoc.getFirstSection().getBody().getFirstParagraph().appendChild(
            new Run(srcDoc, "Source document first paragraph text."));
        dstDoc.getFirstSection().getBody().getFirstParagraph().appendChild(
            new Run(dstDoc, "Destination document first paragraph text."));

        // Every node has a parent document, which is the document that contains the node.
        // Inserting a node into a document that the node does not belong to will throw an exception.
        Assert.assertNotEquals(dstDoc, srcDoc.getFirstSection().getDocument());
        Assert.<IllegalArgumentException>Throws(() => { dstDoc.appendChild(srcDoc.getFirstSection()); });

        // Use the ImportNode method to create a copy of a node, which will have the document
        // that called the ImportNode method set as its new owner document.
        Section importedSection = (Section)dstDoc.importNode(srcDoc.getFirstSection(), true);

        Assert.assertEquals(dstDoc, importedSection.getDocument());

        // We can now insert the node into the document.
        dstDoc.appendChild(importedSection);

        Assert.assertEquals("Destination document first paragraph text.\r\nSource document first paragraph text.\r\n",
            dstDoc.toString(SaveFormat.TEXT));
        //ExEnd

        Assert.assertNotEquals(importedSection, srcDoc.getFirstSection());
        Assert.assertNotEquals(importedSection.getDocument(), srcDoc.getFirstSection().getDocument());
        Assert.assertEquals(importedSection.getBody().getFirstParagraph().getText(),
            srcDoc.getFirstSection().getBody().getFirstParagraph().getText());
    }

    @Test
    public void importNodeCustom() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase.ImportNode(Node, System.Boolean, ImportFormatMode)
        //ExSummary:Shows how to import node from source document to destination document with specific options.
        // Create two documents and add a character style to each document.
        // Configure the styles to have the same name, but different text formatting.
        Document srcDoc = new Document();
        Style srcStyle = srcDoc.getStyles().add(StyleType.CHARACTER, "My style");
        srcStyle.getFont().setName("Courier New");
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.getFont().setStyle(srcStyle);
        srcBuilder.writeln("Source document text.");

        Document dstDoc = new Document();
        Style dstStyle = dstDoc.getStyles().add(StyleType.CHARACTER, "My style");
        dstStyle.getFont().setName("Calibri");
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.getFont().setStyle(dstStyle);
        dstBuilder.writeln("Destination document text.");

        // Import the Section from the destination document into the source document, causing a style name collision.
        // If we use destination styles, then the imported source text with the same style name
        // as destination text will adopt the destination style.
        Section importedSection = (Section)dstDoc.importNode(srcDoc.getFirstSection(), true, ImportFormatMode.USE_DESTINATION_STYLES);
        Assert.assertEquals("Source document text.", importedSection.getBody().getParagraphs().get(0).getRuns().get(0).getText().trim()); //ExSkip
        Assert.assertNull(dstDoc.getStyles().get("My style_0")); //ExSkip
        Assert.assertEquals(dstStyle.getFont().getName(), importedSection.getBody().getFirstParagraph().getRuns().get(0).getFont().getName());
        Assert.assertEquals(dstStyle.getName(), importedSection.getBody().getFirstParagraph().getRuns().get(0).getFont().getStyleName());

        // If we use ImportFormatMode.KeepDifferentStyles, the source style is preserved,
        // and the naming clash resolves by adding a suffix.
        dstDoc.importNode(srcDoc.getFirstSection(), true, ImportFormatMode.KEEP_DIFFERENT_STYLES);
        Assert.assertEquals(dstStyle.getFont().getName(), dstDoc.getStyles().get("My style").getFont().getName());
        Assert.assertEquals(srcStyle.getFont().getName(), dstDoc.getStyles().get("My style_0").getFont().getName());
        //ExEnd
    }

    @Test
    public void backgroundShape() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase.BackgroundShape
        //ExSummary:Shows how to set a background shape for every page of a document.
        Document doc = new Document();

        Assert.assertNull(doc.getBackgroundShape());

        // The only shape type that we can use as a background is a rectangle.
        Shape shapeRectangle = new Shape(doc, ShapeType.RECTANGLE);

        // There are two ways of using this shape as a page background.
        // 1 -  A flat color:
        shapeRectangle.setFillColor(java.awt.Color.LightBlue);
        doc.setBackgroundShape(shapeRectangle);

        doc.save(getArtifactsDir() + "DocumentBase.BackgroundShape.FlatColor.docx");

        // 2 -  An image:
        shapeRectangle = new Shape(doc, ShapeType.RECTANGLE);
        shapeRectangle.getImageData().setImage(getImageDir() + "Transparent background logo.png");

        // Adjust the image's appearance to make it more suitable as a watermark.
        shapeRectangle.getImageData().setContrast(0.2);
        shapeRectangle.getImageData().setBrightness(0.7);

        doc.setBackgroundShape(shapeRectangle);

        Assert.assertTrue(doc.getBackgroundShape().hasImage());

        // Microsoft Word does not support shapes with images as backgrounds,
        // but we can still see these backgrounds in other save formats such as .pdf.
        doc.save(getArtifactsDir() + "DocumentBase.BackgroundShape.Image.pdf");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBase.BackgroundShape.FlatColor.docx");

        Assert.assertEquals(java.awt.Color.LightBlue.getRGB(), doc.getBackgroundShape().getFillColor().getRGB());
        Assert.<IllegalArgumentException>Throws(() =>
        {
            doc.setBackgroundShape(new Shape(doc, ShapeType.TRIANGLE));
        });

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(getArtifactsDir() + "DocumentBase.BackgroundShape.Image.pdf");
        XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

        Assert.AreEqual(400, pdfDocImage.Width);
        Assert.AreEqual(400, pdfDocImage.Height);
        Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
    }

    //ExStart
    //ExFor:DocumentBase.ResourceLoadingCallback
    //ExFor:IResourceLoadingCallback
    //ExFor:IResourceLoadingCallback.ResourceLoading(ResourceLoadingArgs)
    //ExFor:ResourceLoadingAction
    //ExFor:ResourceLoadingArgs
    //ExFor:ResourceLoadingArgs.OriginalUri
    //ExFor:ResourceLoadingArgs.ResourceType
    //ExFor:ResourceLoadingArgs.SetData(Byte[])
    //ExFor:ResourceType
    //ExSummary:Shows how to customize the process of loading external resources into a document.
    @Test //ExSkip
    public void resourceLoadingCallback() throws Exception
    {
        Document doc = new Document();
        doc.setResourceLoadingCallback(new ImageNameHandler());

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Images usually are inserted using a URI, or a byte array.
        // Every instance of a resource load will call our callback's ResourceLoading method.
        builder.insertImage("Google logo");
        builder.insertImage("Aspose logo");
        builder.insertImage("Watermark");

        Assert.assertEquals(3, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        doc.save(getArtifactsDir() + "DocumentBase.ResourceLoadingCallback.docx");
        testResourceLoadingCallback(new Document(getArtifactsDir() + "DocumentBase.ResourceLoadingCallback.docx")); //ExSkip
    }

    /// <summary>
    /// Allows us to load images into a document using predefined shorthands, as opposed to URIs.
    /// This will separate image loading logic from the rest of the document construction.
    /// </summary>
    private static class ImageNameHandler implements IResourceLoadingCallback
    {
        public /*ResourceLoadingAction*/int resourceLoading(ResourceLoadingArgs args) throws Exception
        {
            // If this callback encounters one of the image shorthands while loading an image,
            // it will apply unique logic for each defined shorthand instead of treating it as a URI.
            if (args.getResourceType() == ResourceType.IMAGE)
                switch (gStringSwitchMap.of(args.getOriginalUri()))
                {
                    case /*"Google logo"*/0:
                        WebClient webClient = new WebClient();
                        try /*JAVA: was using*/
                        {
                            args.setData(webClient.DownloadData("http://www.google.com/images/logos/ps_logo2.png"));
                        }
                        finally { if (webClient != null) webClient.close(); }

                        return ResourceLoadingAction.USER_PROVIDED;

                    case /*"Aspose logo"*/1:
                        args.setData(File.readAllBytes(getImageDir() + "Logo.jpg"));

                        return ResourceLoadingAction.USER_PROVIDED;

                    case /*"Watermark"*/2:
                        args.setData(File.readAllBytes(getImageDir() + "Transparent background logo.png"));

                        return ResourceLoadingAction.USER_PROVIDED;
                }

            return ResourceLoadingAction.DEFAULT;
        }
    }
    //ExEnd

    private void testResourceLoadingCallback(Document doc) throws Exception
    {
        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true))
        {
            Assert.assertTrue(shape.hasImage());
            Assert.IsNotEmpty(shape.getImageData().getImageBytes());
        }

        TestUtil.verifyWebResponseStatusCode(HttpStatusCode.OK, "http://www.google.com/images/logos/ps_logo2.png");
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"Google logo",
		"Aspose logo",
		"Watermark"
	);

}
