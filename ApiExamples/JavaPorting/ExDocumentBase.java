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
import com.aspose.words.GlossaryDocument;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import org.testng.Assert;
import com.aspose.words.Run;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.Section;
import com.aspose.words.SaveFormat;
import com.aspose.words.Style;
import com.aspose.words.StyleType;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ImportFormatMode;
import com.aspose.ms.System.msString;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.NodeType;
import com.aspose.words.IResourceLoadingCallback;
import com.aspose.words.ResourceLoadingAction;
import com.aspose.words.ResourceLoadingArgs;
import com.aspose.words.ResourceType;
import java.awt.image.BufferedImage;
import com.aspose.BitmapPal;


@Test
public class ExDocumentBase extends ApiExampleBase
{
    @Test
    public void constructor() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase
        //ExSummary:Shows how to initialize the subclasses of DocumentBase.
        // DocumentBase is the abstract base class for the Document and GlossaryDocument classes
        Document doc = new Document();

        GlossaryDocument glossaryDoc = new GlossaryDocument();
        doc.setGlossaryDocument(glossaryDoc);
        //ExEnd
    }

    @Test
    public void setPageColor() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase.PageColor
        //ExSummary:Shows how to set the page color.
        Document doc = new Document();

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
        //ExSummary:Shows how to import node from source document to destination document.
        Document src = new Document();
        Document dst = new Document();

        // Add text to both documents
        src.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(src, "Source document first paragraph text."));
        dst.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(dst, "Destination document first paragraph text."));

        // In order for a child node to be successfully appended to another node in a document,
        // both nodes must have the same parent document, or an exception is thrown
        msAssert.areNotEqual(dst, src.getFirstSection().getDocument());
        Assert.<IllegalArgumentException>Throws(() => { dst.appendChild(src.getFirstSection()); });

        // For that reason, we can't just append a section of the source document to the destination document using Node.AppendChild()
        // Document.ImportNode() lets us get around this by creating a clone of a node and sets its parent to the calling document
        Section importedSection = (Section)dst.importNode(src.getFirstSection(), true);

        // Now it is ready to be placed in the document
        dst.appendChild(importedSection);

        // Our document now contains both the original and imported section
        Assert.assertEquals("Destination document first paragraph text.\r\nSource document first paragraph text.\r\n",
            dst.toString(SaveFormat.TEXT));
        //ExEnd

        msAssert.areNotEqual(importedSection, src.getFirstSection());
        msAssert.areNotEqual(importedSection.getDocument(), src.getFirstSection().getDocument());
        Assert.assertEquals(importedSection.getBody().getFirstParagraph().getText(),
            src.getFirstSection().getBody().getFirstParagraph().getText());
    }

    @Test
    public void importNodeCustom() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase.ImportNode(Node, System.Boolean, ImportFormatMode)
        //ExSummary:Shows how to import node from source document to destination document with specific options.
        // Create two documents with two styles that differ in font but have the same name
        Document src = new Document();
        Style srcStyle = src.getStyles().add(StyleType.CHARACTER, "My style");
        srcStyle.getFont().setName("Courier New");
        DocumentBuilder srcBuilder = new DocumentBuilder(src);
        srcBuilder.getFont().setStyle(srcStyle);
        srcBuilder.writeln("Source document text.");

        Document dst = new Document();
        Style dstStyle = dst.getStyles().add(StyleType.CHARACTER, "My style");
        dstStyle.getFont().setName("Calibri");
        DocumentBuilder dstBuilder = new DocumentBuilder(dst);
        dstBuilder.getFont().setStyle(dstStyle);
        dstBuilder.writeln("Destination document text.");

        // Import the Section from the destination document into the source document, causing a style name collision
        // If we use destination styles then the imported source text with the same style name as destination text
        // will adopt the destination style 
        Section importedSection = (Section)dst.importNode(src.getFirstSection(), true, ImportFormatMode.USE_DESTINATION_STYLES);
        Assert.assertEquals("Source document text.", msString.trim(importedSection.getBody().getParagraphs().get(0).getRuns().get(0).getText())); //ExSkip
        Assert.assertNull(dst.getStyles().get("My style_0")); //ExSkip
        Assert.assertEquals(dstStyle.getFont().getName(), importedSection.getBody().getFirstParagraph().getRuns().get(0).getFont().getName());
        Assert.assertEquals(dstStyle.getName(), importedSection.getBody().getFirstParagraph().getRuns().get(0).getFont().getStyleName());

        // If we use ImportFormatMode.KeepDifferentStyles,
        // the source style is preserved and the naming clash is resolved by adding a suffix 
        dst.importNode(src.getFirstSection(), true, ImportFormatMode.KEEP_DIFFERENT_STYLES);
        Assert.assertEquals(dstStyle.getFont().getName(), dst.getStyles().get("My style").getFont().getName());
        Assert.assertEquals(srcStyle.getFont().getName(), dst.getStyles().get("My style_0").getFont().getName());
        //ExEnd
    }

    @Test
    public void backgroundShape() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase.BackgroundShape
        //ExSummary:Shows how to set the background shape of a document.
        Document doc = new Document();
        Assert.assertNull(doc.getBackgroundShape());

        // A background shape can only be a rectangle
        // We will set the color of this rectangle to light blue
        Shape shapeRectangle = new Shape(doc, ShapeType.RECTANGLE);
        doc.setBackgroundShape(shapeRectangle);

        // This rectangle covers the entire page in the output document
        // We can also do this by setting doc.PageColor
        shapeRectangle.setFillColor(java.awt.Color.LightBlue);
        doc.save(getArtifactsDir() + "DocumentBase.BackgroundShapeFlatColor.docx");

        // Setting the image will override the flat background color with the image
        shapeRectangle.getImageData().setImage(getImageDir() + "Transparent background logo.png");
        Assert.assertTrue(doc.getBackgroundShape().hasImage());

        // This image is a photo with a white background
        // To make it suitable as a watermark, we will need to do some image processing
        // The default values for these variables are 0.5, so here we are lowering the contrast and increasing the brightness
        shapeRectangle.getImageData().setContrast(0.2);
        shapeRectangle.getImageData().setBrightness(0.7);

        // Microsoft Word does not support images in background shapes, so even though we set the background as an image,
        // the output will show a light blue background like before
        // However, we can see our watermark in an output pdf
        doc.save(getArtifactsDir() + "DocumentBase.BackgroundShape.pdf");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBase.BackgroundShapeFlatColor.docx");
        Assert.assertEquals(java.awt.Color.LightBlue.getRGB(), doc.getBackgroundShape().getFillColor().getRGB());
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
    //ExSummary:Shows how to process inserted resources differently.
    @Test //ExSkip
    public void resourceLoadingCallback() throws Exception
    {
        Document doc = new Document();

        // Enable our custom image loading
        doc.setResourceLoadingCallback(new ImageNameHandler());

        DocumentBuilder builder = new DocumentBuilder(doc);

        // We usually insert images as a uri or byte array, but there are many other possibilities with ResourceLoadingCallback
        // In this case we are referencing images with simple names and keep the image fetching logic somewhere else
        builder.insertImage("Google Logo");
        builder.insertImage("Aspose Logo");
        builder.insertImage("My Watermark");

        // Images belong to Shape objects, which are placed and scaled in the document
        Assert.assertEquals(3, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        doc.save(getArtifactsDir() + "DocumentBase.ResourceLoadingCallback.docx");
        testResourceLoadingCallback(new Document(getArtifactsDir() + "DocumentBase.ResourceLoadingCallback.docx")); //ExSkip
    }

    private static class ImageNameHandler implements IResourceLoadingCallback
    {
        public /*ResourceLoadingAction*/int resourceLoading(ResourceLoadingArgs args)
        {
            if (args.getResourceType() == ResourceType.IMAGE)
            {
                // builder.InsertImage expects a uri so inputs like "Google Logo" would normally trigger a FileNotFoundException
                // We can still process those inputs and find an image any way we like, as long as an image byte array is passed to args.SetData()
                if ("Google Logo".equals(args.getOriginalUri()))
                {
                    WebClient webClient = new WebClient();
                    try /*JAVA: was using*/
                    {
                        byte[] imageBytes =
                            webClient.DownloadData("http://www.google.com/images/logos/ps_logo2.png");
                        args.setData(imageBytes);
                        // We need this return statement any time a resource is loaded in a custom manner
                        return ResourceLoadingAction.USER_PROVIDED;
                    }
                    finally { if (webClient != null) webClient.close(); }
                }

                if ("Aspose Logo".equals(args.getOriginalUri()))
                {
                    WebClient webClient = new WebClient();
                    try /*JAVA: was using*/
                    {
                        byte[] imageBytes = webClient.DownloadData(getAsposeLogoUrl());
                        args.setData(imageBytes);
                        return ResourceLoadingAction.USER_PROVIDED;
                    }
                    finally { if (webClient != null) webClient.close(); }
                }

                // We can find and add an image any way we like, as long as args.SetData() is called with some image byte array as a parameter
                if ("My Watermark".equals(args.getOriginalUri()))
                {
                    BufferedImage watermark = BitmapPal.loadNativeImage(getImageDir() + "Transparent background logo.png");

                    ImageConverter converter = new ImageConverter();
                    byte[] imageBytes = (byte[])converter.ConvertTo(watermark, byte[].class);
                    args.setData(imageBytes);

                    return ResourceLoadingAction.USER_PROVIDED;
                }
            }

            // All other resources such as documents, CSS stylesheets and images passed as uris are handled as they were normally
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
    }
    }
