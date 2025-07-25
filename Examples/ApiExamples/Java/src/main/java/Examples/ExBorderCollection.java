package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.util.Iterator;

public class ExBorderCollection extends ApiExampleBase {
    @Test
    public void getBordersEnumerator() throws Exception {
        //ExStart
        //ExFor:BorderCollection.GetEnumerator
        //ExSummary:Shows how to iterate over and edit all of the borders in a paragraph format object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure the builder's paragraph format settings to create a green wave border on all sides.
        BorderCollection borders = builder.getParagraphFormat().getBorders();

        Iterator<Border> enumerator = borders.iterator();
        while (enumerator.hasNext()) {
            Border border = enumerator.next();
            border.setColor(Color.green);
            border.setLineStyle(LineStyle.WAVE);
            border.setLineWidth(3.0);
        }

        // Insert a paragraph. Our border settings will determine the appearance of its border.
        builder.writeln("Hello world!");

        doc.save(getArtifactsDir() + "BorderCollection.GetBordersEnumerator.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "BorderCollection.GetBordersEnumerator.docx");

        for (Border border : doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBorders()) {
            Assert.assertEquals(Color.green.getRGB(), border.getColor().getRGB());
            Assert.assertEquals(LineStyle.WAVE, border.getLineStyle());
            Assert.assertEquals(3.0d, border.getLineWidth());
        }
    }

    @Test
    public void removeAllBorders() throws Exception {
        //ExStart
        //ExFor:BorderCollection.ClearFormatting
        //ExSummary:Shows how to remove all borders from all paragraphs in a document.
        Document doc = new Document(getMyDir() + "Borders.docx");

        // The first paragraph of this document has visible borders with these settings.
        BorderCollection firstParagraphBorders = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBorders();

        Assert.assertEquals(Color.RED.getRGB(), firstParagraphBorders.getColor().getRGB());
        Assert.assertEquals(LineStyle.SINGLE, firstParagraphBorders.getLineStyle());
        Assert.assertEquals(3.0d, firstParagraphBorders.getLineWidth());

        // Use the "ClearFormatting" method on each paragraph to remove all borders.
        for (Paragraph paragraph : doc.getFirstSection().getBody().getParagraphs()) {
            paragraph.getParagraphFormat().getBorders().clearFormatting();

            for (Border border : paragraph.getParagraphFormat().getBorders()) {
                Assert.assertEquals(0, border.getColor().getRGB());
                Assert.assertEquals(LineStyle.NONE, border.getLineStyle());
                Assert.assertEquals(0.0d, border.getLineWidth());
            }
        }

        doc.save(getArtifactsDir() + "BorderCollection.RemoveAllBorders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "BorderCollection.RemoveAllBorders.docx");

        for (Border border : doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBorders()) {
            Assert.assertEquals(0, border.getColor().getRGB());
            Assert.assertEquals(LineStyle.NONE, border.getLineStyle());
            Assert.assertEquals(0.0d, border.getLineWidth());
        }
    }
}
