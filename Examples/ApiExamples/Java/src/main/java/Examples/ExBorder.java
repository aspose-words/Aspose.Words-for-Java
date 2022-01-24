package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.text.MessageFormat;

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
        //ExFor:Border.IsVisible
        //ExSummary:Shows how to remove borders from a paragraph.
        Document doc = new Document(getMyDir() + "Borders.docx");

        // Each paragraph has an individual set of borders.
        // We can access the settings for the appearance of these borders via the paragraph format object.
        BorderCollection borders = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBorders();

        Assert.assertEquals(Color.RED.getRGB(), borders.get(0).getColor().getRGB());
        Assert.assertEquals(3.0d, borders.get(0).getLineWidth());
        Assert.assertEquals(LineStyle.SINGLE, borders.get(0).getLineStyle());
        Assert.assertTrue(borders.get(0).isVisible());

        // We can remove a border at once by running the ClearFormatting method. 
        // Running this method on every border of a paragraph will remove all its borders.
        for (Border border : borders)
            border.clearFormatting();

        Assert.assertEquals(0, borders.get(0).getColor().getRGB());
        Assert.assertEquals(0.0d, borders.get(0).getLineWidth());
        Assert.assertEquals(LineStyle.NONE, borders.get(0).getLineStyle());
        Assert.assertFalse(borders.get(0).isVisible());

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
    public void sharedElements() throws Exception {
        //ExStart
        //ExFor:Border.Equals(Object)
        //ExFor:Border.Equals(Border)
        //ExFor:Border.GetHashCode
        //ExFor:BorderCollection.Count
        //ExFor:BorderCollection.Equals(BorderCollection)
        //ExFor:BorderCollection.Item(Int32)
        //ExSummary:Shows how border collections can share elements.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Paragraph 1.");
        builder.write("Paragraph 2.");

        // Since we used the same border configuration while creating
        // these paragraphs, their border collections share the same elements.
        BorderCollection firstParagraphBorders = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBorders();
        BorderCollection secondParagraphBorders = builder.getCurrentParagraph().getParagraphFormat().getBorders();
        Assert.assertEquals(6, firstParagraphBorders.getCount()); //ExSkip

        for (int i = 0; i < firstParagraphBorders.getCount(); i++) {
            Assert.assertTrue(firstParagraphBorders.get(i).equals(secondParagraphBorders.get(i)));
            Assert.assertEquals(firstParagraphBorders.get(i).hashCode(), secondParagraphBorders.get(i).hashCode());
            Assert.assertFalse(firstParagraphBorders.get(i).isVisible());
        }

        for (Border border : secondParagraphBorders)
            border.setLineStyle(LineStyle.DOT_DASH);

        // After changing the line style of the borders in just the second paragraph,
        // the border collections no longer share the same elements.
        for (int i = 0; i < firstParagraphBorders.getCount(); i++) {
            Assert.assertFalse(firstParagraphBorders.get(i).equals(secondParagraphBorders.get(i)));
            Assert.assertNotEquals(firstParagraphBorders.get(i).hashCode(), secondParagraphBorders.get(i).hashCode());

            // Changing the appearance of an empty border makes it visible.
            Assert.assertTrue(secondParagraphBorders.get(i).isVisible());
        }

        doc.save(getArtifactsDir() + "Border.SharedElements.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Border.SharedElements.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        for (Border testBorder : paragraphs.get(0).getParagraphFormat().getBorders())
            Assert.assertEquals(LineStyle.NONE, testBorder.getLineStyle());

        for (Border testBorder : paragraphs.get(1).getParagraphFormat().getBorders())
            Assert.assertEquals(LineStyle.DOT_DASH, testBorder.getLineStyle());
    }

    @Test
    public void horizontalBorders() throws Exception {
        //ExStart
        //ExFor:BorderCollection.Horizontal
        //ExSummary:Shows how to apply settings to horizontal borders to a paragraph's format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a red horizontal border for the paragraph. Any paragraphs created afterwards will inherit these border settings.
        BorderCollection borders = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getBorders();
        borders.getHorizontal().setColor(Color.RED);
        borders.getHorizontal().setLineStyle(LineStyle.DASH_SMALL_GAP);
        borders.getHorizontal().setLineWidth(3.0);

        // Write text to the document without creating a new paragraph afterward.
        // Since there is no paragraph underneath, the horizontal border will not be visible.
        builder.write("Paragraph above horizontal border.");

        // Once we add a second paragraph, the border of the first paragraph will become visible.
        builder.insertParagraph();
        builder.write("Paragraph below horizontal border.");

        doc.save(getArtifactsDir() + "Border.HorizontalBorders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Border.HorizontalBorders.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(LineStyle.DASH_SMALL_GAP, paragraphs.get(0).getParagraphFormat().getBorders().getByBorderType(BorderType.HORIZONTAL).getLineStyle());
        Assert.assertEquals(LineStyle.DASH_SMALL_GAP, paragraphs.get(1).getParagraphFormat().getBorders().getByBorderType(BorderType.HORIZONTAL).getLineStyle());
    }

    @Test
    public void verticalBorders() throws Exception {
        //ExStart
        //ExFor:BorderCollection.Horizontal
        //ExFor:BorderCollection.Vertical
        //ExFor:Cell.LastParagraph
        //ExSummary:Shows how to apply settings to vertical borders to a table row's format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a table with red and blue inner borders.
        Table table = builder.startTable();

        for (int i = 0; i < 3; i++) {
            builder.insertCell();
            builder.write(MessageFormat.format("Row {0}, Column 1", i + 1));
            builder.insertCell();
            builder.write(MessageFormat.format("Row {0}, Column 2", i + 1));

            Row row = builder.endRow();
            BorderCollection borders = row.getRowFormat().getBorders();

            // Adjust the appearance of borders that will appear between rows.
            borders.getHorizontal().setColor(Color.RED);
            borders.getHorizontal().setLineStyle(LineStyle.DOT);
            borders.getHorizontal().setLineWidth(2.0d);

            // Adjust the appearance of borders that will appear between cells.
            borders.getVertical().setColor(Color.BLUE);
            borders.getVertical().setLineStyle(LineStyle.DOT);
            borders.getVertical().setLineWidth(2.0d);
        }

        // A row format, and a cell's inner paragraph use different border settings.
        Border border = table.getFirstRow().getFirstCell().getLastParagraph().getParagraphFormat().getBorders().getVertical();

        Assert.assertEquals(0, border.getColor().getRGB());
        Assert.assertEquals(0.0d, border.getLineWidth());
        Assert.assertEquals(LineStyle.NONE, border.getLineStyle());

        doc.save(getArtifactsDir() + "Border.VerticalBorders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Border.VerticalBorders.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        for (Row row : (Iterable<Row>) table.getChildNodes(NodeType.ROW, true)) {
            Assert.assertEquals(Color.RED.getRGB(), row.getRowFormat().getBorders().getHorizontal().getColor().getRGB());
            Assert.assertEquals(LineStyle.DOT, row.getRowFormat().getBorders().getHorizontal().getLineStyle());
            Assert.assertEquals(2.0d, row.getRowFormat().getBorders().getHorizontal().getLineWidth());

            Assert.assertEquals(Color.BLUE.getRGB(), row.getRowFormat().getBorders().getVertical().getColor().getRGB());
            Assert.assertEquals(LineStyle.DOT, row.getRowFormat().getBorders().getVertical().getLineStyle());
            Assert.assertEquals(2.0d, row.getRowFormat().getBorders().getVertical().getLineWidth());
        }
    }
}
