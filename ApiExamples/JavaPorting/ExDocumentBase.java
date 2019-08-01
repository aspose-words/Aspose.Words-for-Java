package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.GlossaryDocument;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.Run;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.Section;
import com.aspose.words.SaveFormat;
import com.aspose.words.Style;
import com.aspose.words.StyleType;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ImportFormatMode;
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
        dst.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(dst,
            "Destination document first paragraph text."));

        // If we want to add the section from doc2 to doc1, we can't just append them like this:
        // dst.AppendChild(src.FirstSection);
        // Uncommenting that line throws an exception because doc2's first section belongs to doc2,
        // but each node in a document must belong to the document
        msAssert.areNotEqual(dst, src.getFirstSection().getDocument());

        // We can create a new node that belongs to the destination document
        Section importedSection = (Section) dst.importNode(src.getFirstSection(), true);

        // It has the same content but it is not the same node nor do they have the same owner
        msAssert.areNotEqual(importedSection, src.getFirstSection());
        msAssert.areNotEqual(importedSection.getDocument(), src.getFirstSection().getDocument());
        msAssert.areEqual(importedSection.getBody().getFirstParagraph().getText(),
            src.getFirstSection().getBody().getFirstParagraph().getText());

        // Now it is ready to be placed in the document
        dst.appendChild(importedSection);

        // Our document does indeed contain both the original and imported section
        msAssert.areEqual("Destination document first paragraph text.\r\nSource document first paragraph text.\r\n",
            dst.toString(SaveFormat.TEXT));
        //ExEnd
    }

    @Test
    public void importNodeCustom() throws Exception
    {
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
    public void backgroundShape() throws Exception
    {
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
        shapeRectangle.setFillColor(java.awt.Color.LightBlue);
        doc.save(getArtifactsDir() + "DocumentBase.BackgroundShapeFlatColor.docx");

        // Setting the image will override the flat background colour with the image
        shapeRectangle.getImageData().setImage(getMyDir() + "Images/Watermark.png");
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
    public void resourceLoadingCallback() throws Exception
    {
        Document doc = new Document();

        // Images belong to NodeType.Shape
        // There are none in a blank document
        msAssert.areEqual(0, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        // Enable our custom image loading
        doc.setResourceLoadingCallback(new ImageNameHandler());

        DocumentBuilder builder = new DocumentBuilder(doc);

        // We usually insert images as a uri or byte array, but there are many other possibilities with ResourceLoadingCallback
        // In this case we are referencing images with simple names and keep the image fetching logic somewhere else
        builder.insertImage("Google Logo");
        builder.insertImage("Aspose Logo");
        builder.insertImage("My Watermark");

        msAssert.areEqual(3, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        doc.save(getArtifactsDir() + "DocumentBase.ResourceLoadingCallback.docx");            
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
                    BufferedImage watermark = BitmapPal.loadNativeImage(getMyDir() + "Images/Watermark.png");

                    ImageConverter converter = new ImageConverter();
                    byte[] imageBytes = (byte[]) converter.ConvertTo(watermark, byte[].class);
                    args.setData(imageBytes);

                    return ResourceLoadingAction.USER_PROVIDED;
                }
            }

            // All other resources such as documents, CSS stylesheets and images passed as uris are handled as they were normally
            return ResourceLoadingAction.DEFAULT;
        }
    }
    //ExEnd
}
