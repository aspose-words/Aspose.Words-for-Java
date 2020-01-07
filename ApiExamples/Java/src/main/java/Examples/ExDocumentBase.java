package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.Shape;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.URISyntaxException;

@Test
public class ExDocumentBase extends ApiExampleBase {
    @Test
    public void constructor() throws Exception {
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
    public void setPageColor() throws Exception {
        //ExStart
        //ExFor:DocumentBase.PageColor
        //ExSummary:Shows how to set the page color.
        Document doc = new Document();

        doc.setPageColor(Color.lightGray);

        doc.save(getArtifactsDir() + "DocumentBase.SetPageColor.docx");
        //ExEnd
    }

    @Test
    public void importNode() throws Exception {
        //ExStart
        //ExFor:DocumentBase.ImportNode(Node, Boolean)
        //ExSummary:Shows how to import node from source document to destination document.
        Document src = new Document();
        Document dst = new Document();

        // Add text to both documents
        src.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(src, "Source document first paragraph text."));
        dst.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(dst,
                "Destination document first paragraph text."));

        // If we want to add the section from doc2 to doc1, we can't just append them like this:
        // dst.AppendChild(src.FirstSection);
        // Uncommenting that line throws an exception because doc2's first section belongs to doc2,
        // but each node in a document must belong to the document
        Assert.assertNotEquals(src.getFirstSection().getDocument(), dst);

        // We can create a new node that belongs to the destination document
        Section importedSection = (Section) dst.importNode(src.getFirstSection(), true);

        // It has the same content but it is not the same node nor do they have the same owner
        Assert.assertNotEquals(src.getFirstSection(), importedSection);
        Assert.assertNotEquals(src.getFirstSection().getDocument(), importedSection.getDocument());
        Assert.assertEquals(src.getFirstSection().getBody().getFirstParagraph().getText(),
                importedSection.getBody().getFirstParagraph().getText());

        // Now it is ready to be placed in the document
        dst.appendChild(importedSection);

        // Our document does indeed contain both the original and imported section
        Assert.assertEquals("Destination document first paragraph text.\r\nSource document first paragraph text.\r\n",
                dst.toString(SaveFormat.TEXT));
        //ExEnd
    }

    @Test
    public void importNodeCustom() throws Exception {
        //ExStart
        //ExFor:DocumentBase.ImportNode(Node, System.Boolean, ImportFormatMode)
        //ExSummary:Shows how to import node from source document to destination document with specific options.
        // Create two documents with two styles that aren't the same but have the same name
        Document src = new Document();
        Style srcStyle = src.getStyles().add(StyleType.CHARACTER, "My style");
        DocumentBuilder srcBuilder = new DocumentBuilder(src);
        srcBuilder.getFont().setStyle(srcStyle);
        srcBuilder.writeln("Source document text.");

        Document dst = new Document();
        Style dstStyle = dst.getStyles().add(StyleType.CHARACTER, "My style");
        dstStyle.getFont().setBold(true);
        DocumentBuilder dstBuilder = new DocumentBuilder(dst);
        dstBuilder.getFont().setStyle(dstStyle);
        srcBuilder.writeln("Destination document text.");

        dst.importNode(src.getFirstSection(), true, ImportFormatMode.USE_DESTINATION_STYLES);

        Assert.assertNull(dst.getStyles().get("My style_0"));

        dst.importNode(src.getFirstSection(), true, ImportFormatMode.KEEP_DIFFERENT_STYLES);

        Assert.assertNotNull(dst.getStyles().get("My style_0"));
        //ExEnd
    }

    @Test
    public void backgroundShape() throws Exception {
        //ExStart
        //ExFor:DocumentBase.BackgroundShape
        //ExSummary:Shows how to set the background shape of a document.
        Document doc = new Document();
        Assert.assertNull(doc.getBackgroundShape());

        // A background shape can only be a rectangle
        // We will set the colour of this rectangle to light blue
        Shape shapeRectangle = new Shape(doc, ShapeType.RECTANGLE);
        doc.setBackgroundShape(shapeRectangle);

        // This rectangle covers the entire page in the output document
        // We can also do this by setting doc.PageColor
        shapeRectangle.setFillColor(Color.blue);
        doc.save(getArtifactsDir() + "DocumentBase.BackgroundShapeFlatColor.docx");

        // Setting the image will override the flat background colour with the image
        shapeRectangle.getImageData().setImage(getImageDir() + "Watermark.png");
        Assert.assertTrue(doc.getBackgroundShape().hasImage());

        // This image is a photo with a white background
        // To make it suitable as a watermark, we will need to do some image processing
        // The default values for these variables are 0.5, so here we are lowering the contrast and increasing the brightness
        shapeRectangle.getImageData().setContrast(0.2);
        shapeRectangle.getImageData().setBrightness(0.7);

        // Microsoft Word does not support images in background shapes, so even though we set the background as an image,
        // the output will show a light blue background like before
        // However, we can see our watermark in an output pdf
        doc.save(getArtifactsDir() + "DocumentBase.BackgroundShapeWatermark.pdf");
        //ExEnd
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
    public void resourceLoadingCallback() throws Exception {
        Document doc = new Document();

        // Images belong to NodeType.Shape
        // There are none in a blank document
        Assert.assertEquals(doc.getChildNodes(NodeType.SHAPE, true).getCount(), 0);

        // Enable our custom image loading
        doc.setResourceLoadingCallback(new ImageNameHandler());

        DocumentBuilder builder = new DocumentBuilder(doc);

        // We usually insert images as a uri or byte array, but there are many other possibilities with ResourceLoadingCallback
        // In this case we are referencing images with simple names and keep the image fetching logic somewhere else
        builder.insertImage("Google Logo");
        builder.insertImage("Aspose Logo");
        builder.insertImage("My Watermark");

        Assert.assertEquals(doc.getChildNodes(NodeType.SHAPE, true).getCount(), 3);

        doc.save(getArtifactsDir() + "DocumentBase.ResourceLoadingCallback.docx");
    }

    private static class ImageNameHandler implements IResourceLoadingCallback {
        public int resourceLoading(final ResourceLoadingArgs args) throws URISyntaxException, IOException {
            if (args.getResourceType() == ResourceType.IMAGE) {
                // builder.InsertImage expects a uri so inputs like "Google Logo" would normally trigger a FileNotFoundException
                // We can still process those inputs and find an image any way we like, as long as an image byte array is passed to args.SetData()
                if ("Google Logo".equals(args.getOriginalUri())) {
                    args.setData(DocumentHelper.getBytesFromStream(new URI("http://www.google.com/images/logos/ps_logo2.png").toURL().openStream()));

                    return ResourceLoadingAction.USER_PROVIDED;
                }

                if ("Aspose Logo".equals(args.getOriginalUri())) {
                    args.setData(DocumentHelper.getBytesFromStream(getAsposelogoUri().toURL().openStream()));

                    return ResourceLoadingAction.USER_PROVIDED;
                }

                // We can find and add an image any way we like, as long as args.SetData() is called with some image byte array as a parameter
                if ("My Watermark".equals(args.getOriginalUri())) {
                    InputStream imageStream = new FileInputStream(getImageDir() + "Watermark.png");
                    args.setData(DocumentHelper.getBytesFromStream(imageStream));

                    return ResourceLoadingAction.USER_PROVIDED;
                }
            }

            // All other resources such as documents, CSS stylesheets and images passed as uris are handled as they were normally
            return ResourceLoadingAction.DEFAULT;
        }
    }
    //ExEnd
}
