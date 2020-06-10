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
import org.testng.annotations.Test;

import java.awt.*;

public class ExBorder extends ApiExampleBase {
    @Test
    public void fontBorder() throws Exception {
        //ExStart
        //ExFor:Border
        //ExFor:Border.Color
        //ExFor:Border.LineWidth
        //ExFor:Border.LineStyle
        //ExFor:Font.Border
        //ExFor:LineStyle
        //ExFor:Font
        //ExFor:DocumentBuilder.Font
        //ExFor:DocumentBuilder.Write(String)
        //ExSummary:Shows how to insert a string surrounded by a border into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().getBorder().setColor(Color.GREEN);
        builder.getFont().getBorder().setLineWidth(2.5);
        builder.getFont().getBorder().setLineStyle(LineStyle.DASH_DOT_STROKER);

        builder.write("Text surrounded by green border.");

        doc.save(getArtifactsDir() + "Border.FontBorder.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Border.FontBorder.docx");
        Border border = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).getFont().getBorder();

        Assert.assertEquals(Color.GREEN.getRGB(), border.getColor().getRGB());
        Assert.assertEquals(2.5d, border.getLineWidth());
        Assert.assertEquals(LineStyle.DASH_DOT_STROKER, border.getLineStyle());
    }

    @Test
    public void paragraphTopBorder() throws Exception {
        //ExStart
        //ExFor:BorderCollection
        //ExFor:Border
        //ExFor:BorderType
        //ExFor:ParagraphFormat.Borders
        //ExSummary:Shows how to insert a paragraph with a top border.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Border topBorder = builder.getParagraphFormat().getBorders().getByBorderType(BorderType.TOP);
        topBorder.setColor(Color.RED);
        topBorder.setLineWidth(4.0d);
        topBorder.setLineStyle(LineStyle.DASH_SMALL_GAP);

        builder.writeln("Text with a red top border.");

        doc.save(getArtifactsDir() + "Border.ParagraphTopBorder.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Border.ParagraphTopBorder.docx");
        Border border = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBorders().getByBorderType(BorderType.TOP);

        Assert.assertEquals(Color.RED.getRGB(), border.getColor().getRGB());
        Assert.assertEquals(4.0d, border.getLineWidth());
        Assert.assertEquals(LineStyle.DASH_SMALL_GAP, border.getLineStyle());
    }

    @Test
    public void clearFormatting() throws Exception {
        //ExStart
        //ExFor:Border.ClearFormatting
        //ExSummary:Shows how to remove borders from a paragraph.
        Document doc = new Document(getMyDir() + "Borders.docx");

        // Get the first paragraph's collection of borders
        DocumentBuilder builder = new DocumentBuilder(doc);
        BorderCollection borders = builder.getParagraphFormat().getBorders();
        Assert.assertEquals(Color.RED.getRGB(), borders.get(0).getColor().getRGB()); //ExSkip
        Assert.assertEquals(3.0d, borders.get(0).getLineWidth()); // ExSkip
        Assert.assertEquals(LineStyle.SINGLE, borders.get(0).getLineStyle()); // ExSkip

        for (Border border : borders) {
            border.clearFormatting();
        }

        builder.getCurrentParagraph().getRuns().get(0).setText("Paragraph with no border");

        doc.save(getArtifactsDir() + "Border.ClearFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Border.ClearFormatting.docx");

        for (Border testBorder : doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBorders()) {
            Assert.assertEquals(0, testBorder.getColor().getRGB());
            Assert.assertEquals(0.0d, testBorder.getLineWidth());
            Assert.assertEquals(LineStyle.NONE, testBorder.getLineStyle());
        }
    }

    @Test
    public void equalityCountingAndVisibility() throws Exception {
        //ExStart
        //ExFor:Border.Equals(Object)
        //ExFor:Border.Equals(Border)
        //ExFor:Border.GetHashCode
        //ExFor:Border.IsVisible
        //ExFor:BorderCollection.Count
        //ExFor:BorderCollection.Equals(BorderCollection)
        //ExFor:BorderCollection.Item(Int32)
        //ExSummary:Shows the equality of BorderCollections as well counting, visibility of their elements.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getCurrentParagraph().appendChild(new Run(doc, "Paragraph 1."));

        Paragraph firstParagraph = doc.getFirstSection().getBody().getFirstParagraph();
        BorderCollection firstParaBorders = firstParagraph.getParagraphFormat().getBorders();

        builder.insertParagraph();
        builder.getCurrentParagraph().appendChild(new Run(doc, "Paragraph 2."));

        Paragraph secondParagraph = builder.getCurrentParagraph();
        BorderCollection secondParaBorders = secondParagraph.getParagraphFormat().getBorders();

        // Two paragraphs have two different BorderCollections, but share the elements that are in from the first paragraph
        for (int i = 0; i < firstParaBorders.getCount(); i++) {
            Assert.assertTrue(firstParaBorders.get(i).equals(secondParaBorders.get(i)));
            Assert.assertEquals(firstParaBorders.get(i).hashCode(), secondParaBorders.get(i).hashCode());

            // Borders are invisible by default
            Assert.assertFalse(firstParaBorders.get(i).isVisible());
        }

        // Each border in the second paragraph collection becomes no longer the same as its counterpart from the first paragraph collection
        // Change all the elements in the second collection to make it completely different from the first
        Assert.assertEquals(6, secondParaBorders.getCount()); // ExSkip
        for (Border border : secondParaBorders) {
            border.setLineStyle(LineStyle.DOT_DASH);
        }

        // Now the BorderCollections both have their own elements
        for (int i = 0; i < firstParaBorders.getCount(); i++) {
            Assert.assertFalse(firstParaBorders.get(i).equals(secondParaBorders.get(i)));
            Assert.assertNotEquals(firstParaBorders.get(i).hashCode(), secondParaBorders.get(i).hashCode());

            // Changing the line style made the borders visible
            Assert.assertTrue(secondParaBorders.get(i).isVisible());
        }

        doc.save(getArtifactsDir() + "Border.EqualityCountingAndVisibility.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Border.EqualityCountingAndVisibility.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        for (Border testBorder : paragraphs.get(0).getParagraphFormat().getBorders())
            Assert.assertEquals(LineStyle.NONE, testBorder.getLineStyle());

        for (Border testBorder : paragraphs.get(1).getParagraphFormat().getBorders())
            Assert.assertEquals(LineStyle.DOT_DASH, testBorder.getLineStyle());
    }

    @Test
    public void verticalAndHorizontalBorders() throws Exception {
        //ExStart
        //ExFor:BorderCollection.Horizontal
        //ExFor:BorderCollection.Vertical
        //ExFor:Cell.LastParagraph
        //ExSummary:Shows the difference between the Horizontal and Vertical properties of BorderCollection.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A BorderCollection is one of a Paragraph's formatting properties
        Paragraph paragraph = doc.getFirstSection().getBody().getFirstParagraph();
        BorderCollection paragraphBorders = paragraph.getParagraphFormat().getBorders();

        // paragraphBorders belongs to the first paragraph, but these changes will apply to subsequently created paragraphs
        paragraphBorders.getHorizontal().setColor(Color.RED);
        paragraphBorders.getHorizontal().setLineStyle(LineStyle.DASH_SMALL_GAP);
        paragraphBorders.getHorizontal().setLineWidth(3.0);

        // Horizontal borders only appear under a paragraph if there's another paragraph under it
        // Right now the first paragraph has no borders
        builder.getCurrentParagraph().appendChild(new Run(doc, "Paragraph above horizontal border."));

        // Now the first paragraph will have a red dashed line border under it
        // This new second paragraph can have a border too, but only if we add another paragraph underneath it
        builder.insertParagraph();
        builder.getCurrentParagraph().appendChild(new Run(doc, "Paragraph below horizontal border."));

        // A table makes use of both vertical and horizontal properties of BorderCollection
        // Both these properties can only affect the inner borders of a table
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);

        for (int i = 0; i < 3; i++) {
            Row row = new Row(doc);
            BorderCollection rowBorders = row.getRowFormat().getBorders();

            // Vertical borders are ones between rows in a table
            rowBorders.getHorizontal().setColor(Color.RED);
            rowBorders.getHorizontal().setLineStyle(LineStyle.DOT);
            rowBorders.getHorizontal().setLineWidth(2.0d);

            // Vertical borders are ones between cells in a table
            rowBorders.getVertical().setColor(Color.BLUE);
            rowBorders.getVertical().setLineStyle(LineStyle.DOT);
            rowBorders.getVertical().setLineWidth(2.0d);

            // A blue dotted vertical border will appear between cells
            // A red dotted border will appear between rows
            row.appendChild(new Cell(doc));
            row.getLastCell().appendChild(new Paragraph(doc));
            row.getLastCell().getFirstParagraph().appendChild(new Run(doc, "Vertical border to the right."));

            row.appendChild(new Cell(doc));
            row.getLastCell().appendChild(new Paragraph(doc));
            row.getLastCell().getLastParagraph().appendChild(new Run(doc, "Vertical border to the left."));
            table.appendChild(row);
        }

        doc.save(getArtifactsDir() + "Border.VerticalAndHorizontalBorders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Border.VerticalAndHorizontalBorders.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(LineStyle.DASH_SMALL_GAP, paragraphs.get(0).getParagraphFormat().getBorders().getByBorderType(BorderType.HORIZONTAL).getLineStyle());
        Assert.assertEquals(LineStyle.DASH_SMALL_GAP, paragraphs.get(1).getParagraphFormat().getBorders().getByBorderType(BorderType.HORIZONTAL).getLineStyle());

        Table outTable = (Table) doc.getChild(NodeType.TABLE, 0, true);

        for (Row row : (Iterable<Row>) outTable.getChildNodes(NodeType.ROW, true)) {
            Assert.assertEquals(Color.RED.getRGB(), row.getRowFormat().getBorders().getHorizontal().getColor().getRGB());
            Assert.assertEquals(LineStyle.DOT, row.getRowFormat().getBorders().getHorizontal().getLineStyle());
            Assert.assertEquals(2.0d, row.getRowFormat().getBorders().getHorizontal().getLineWidth());

            Assert.assertEquals(Color.BLUE.getRGB(), row.getRowFormat().getBorders().getVertical().getColor().getRGB());
            Assert.assertEquals(LineStyle.DOT, row.getRowFormat().getBorders().getVertical().getLineStyle());
            Assert.assertEquals(2.0d, row.getRowFormat().getBorders().getVertical().getLineWidth());
        }
    }
}
