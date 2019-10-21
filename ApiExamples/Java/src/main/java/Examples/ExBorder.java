package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
        //ExFor:DocumentBuilder.Write
        //ExSummary:Inserts a String surrounded by a border into a document.
        DocumentBuilder builder = new DocumentBuilder();

        builder.getFont().getBorder().setColor(Color.GREEN);
        builder.getFont().getBorder().setLineWidth(2.5);
        builder.getFont().getBorder().setLineStyle(LineStyle.DASH_DOT_STROKER);

        builder.write("run of text in a green border");
        //ExEnd
    }

    @Test
    public void paragraphTopBorder() throws Exception {
        //ExStart
        //ExFor:BorderCollection
        //ExFor:Border
        //ExFor:BorderType
        //ExFor:ParagraphFormat.Borders
        //ExSummary:Inserts a paragraph with a top border.
        DocumentBuilder builder = new DocumentBuilder();

        Border topBorder = builder.getParagraphFormat().getBorders().getByBorderType(BorderType.TOP);
        topBorder.setColor(Color.RED);
        topBorder.setLineStyle(LineStyle.DASH_SMALL_GAP);
        topBorder.setLineWidth(4);

        builder.writeln("Hello World!");
        //ExEnd
    }

    @Test
    public void clearFormatting() throws Exception {
        //ExStart
        //ExFor:Border.ClearFormatting
        //ExSummary:Shows how to remove borders from a paragraph one by one.
        Document doc = new Document(getMyDir() + "Border.Borders.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        BorderCollection borders = builder.getParagraphFormat().getBorders();

        for (Border border : borders) {
            border.clearFormatting();
        }

        builder.getCurrentParagraph().getRuns().get(0).setText("Paragraph with no border");

        doc.save(getArtifactsDir() + "Border.NoBorder.doc");
        //ExEnd
    }

    @Test
    public void equalityCountingAndVisibility() throws Exception {
        //ExStart
        //ExFor:Border.Equals(System.Object)
        //ExFor:Border.GetHashCode
        //ExFor:Border.IsVisible
        //ExFor:BorderCollection.Count
        //ExFor:BorderCollection.Equals(BorderCollection)
        //ExFor:BorderCollection.Item(System.Int32)
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
        // There are always 6 elements in a border collection, and changing all of them will make the second collection completely different from the first
        secondParaBorders.getByBorderType(BorderType.LEFT).setLineStyle(LineStyle.DOT_DASH);
        secondParaBorders.getByBorderType(BorderType.RIGHT).setLineStyle(LineStyle.DOT_DASH);
        secondParaBorders.getByBorderType(BorderType.TOP).setLineStyle(LineStyle.DOT_DASH);
        secondParaBorders.getByBorderType(BorderType.BOTTOM).setLineStyle(LineStyle.DOT_DASH);
        secondParaBorders.getByBorderType(BorderType.VERTICAL).setLineStyle(LineStyle.DOT_DASH);
        secondParaBorders.getByBorderType(BorderType.HORIZONTAL).setLineStyle(LineStyle.DOT_DASH);

        // Now the BorderCollections both have their own elements
        for (int i = 0; i < firstParaBorders.getCount(); i++) {
            Assert.assertFalse(firstParaBorders.get(i).equals(secondParaBorders.get(i)));
            Assert.assertNotEquals(firstParaBorders.get(i).hashCode(), secondParaBorders.get(i).hashCode());

            // Changing the line style made the borders visible
            Assert.assertTrue(secondParaBorders.get(i).isVisible());
        }
        //ExEnd
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
            rowBorders.getHorizontal().setLineWidth(2.0);

            // Vertical borders are ones between cells in a table
            rowBorders.getVertical().setColor(Color.BLUE);
            rowBorders.getVertical().setLineStyle(LineStyle.DOT);
            rowBorders.getVertical().setLineWidth(2.0);

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

        doc.save(getArtifactsDir() + "Border.HorizontalAndVerticalBorders.docx");
        //ExEnd
    }
}
