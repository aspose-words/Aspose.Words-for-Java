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
        Section importedSection = (Section)dst.importNode(src.getFirstSection(), true);

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
    }

            }
