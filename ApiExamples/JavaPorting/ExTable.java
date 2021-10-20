// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Table;
import com.aspose.words.Row;
import com.aspose.words.Cell;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import org.testng.Assert;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PreferredWidth;
import com.aspose.words.RowFormat;
import com.aspose.words.BorderType;
import java.awt.Color;
import com.aspose.words.CellFormat;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.TableCollection;
import com.aspose.ms.System.msConsole;
import com.aspose.words.RowCollection;
import com.aspose.words.CellCollection;
import com.aspose.words.SaveFormat;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Node;
import com.aspose.words.TableAlignment;
import com.aspose.words.LineStyle;
import com.aspose.words.TextureIndex;
import com.aspose.words.BorderCollection;
import com.aspose.words.HeightRule;
import com.aspose.words.TextOrientation;
import com.aspose.words.Border;
import com.aspose.words.FindReplaceOptions;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.PreferredWidthType;
import com.aspose.words.CellMerge;
import com.aspose.ms.System.Drawing.msPoint;
import com.aspose.ms.System.Drawing.Rectangle;
import com.aspose.words.TextWrapping;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.VerticalAlignment;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.TableStyle;
import com.aspose.words.StyleType;
import com.aspose.words.CellVerticalAlignment;
import com.aspose.words.ConditionalStyleType;
import com.aspose.words.ParagraphAlignment;
import java.util.Iterator;
import com.aspose.words.ConditionalStyle;
import com.aspose.words.TableStyleOptions;
import com.aspose.words.ConditionalStyleCollection;
import org.testng.annotations.DataProvider;


@Test
public class ExTable extends ApiExampleBase
{
    @Test
    public void createTable() throws Exception
    {
        //ExStart
        //ExFor:Table
        //ExFor:Row
        //ExFor:Cell
        //ExFor:Table.#ctor(DocumentBase)
        //ExSummary:Shows how to create a table.
        Document doc = new Document();
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);

        // Tables contain rows, which contain cells, which may have paragraphs
        // with typical elements such as runs, shapes, and even other tables.
        // Calling the "EnsureMinimum" method on a table will ensure that
        // the table has at least one row, cell, and paragraph.
        Row firstRow = new Row(doc);
        table.appendChild(firstRow);

        Cell firstCell = new Cell(doc);
        firstRow.appendChild(firstCell);

        Paragraph paragraph = new Paragraph(doc);
        firstCell.appendChild(paragraph);

        // Add text to the first call in the first row of the table.
        Run run = new Run(doc, "Hello world!");
        paragraph.appendChild(run);

        doc.save(getArtifactsDir() + "Table.CreateTable.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.CreateTable.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(1, table.getRows().getCount());
        Assert.assertEquals(1, table.getFirstRow().getCells().getCount());
        Assert.assertEquals("Hello world!\u0007\u0007", table.getText().trim());
    }

    @Test
    public void padding() throws Exception
    {
        //ExStart
        //ExFor:Table.LeftPadding
        //ExFor:Table.RightPadding
        //ExFor:Table.TopPadding
        //ExFor:Table.BottomPadding
        //ExSummary:Shows how to configure content padding in a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Row 1, cell 1.");
        builder.insertCell();
        builder.write("Row 1, cell 2.");
        builder.endTable();
        
        // For every cell in the table, set the distance between its contents and each of its borders. 
        // This table will maintain the minimum padding distance by wrapping text.
        table.setLeftPadding(30.0);
        table.setRightPadding(60.0);
        table.setTopPadding(10.0);
        table.setBottomPadding(90.0);
        table.setPreferredWidth(PreferredWidth.fromPoints(250.0));

        doc.save(getArtifactsDir() + "DocumentBuilder.SetRowFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilder.SetRowFormatting.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(30.0d, table.getLeftPadding());
        Assert.assertEquals(60.0d, table.getRightPadding());
        Assert.assertEquals(10.0d, table.getTopPadding());
        Assert.assertEquals(90.0d, table.getBottomPadding());
    }

    @Test
    public void rowCellFormat() throws Exception
    {
        //ExStart
        //ExFor:Row.RowFormat
        //ExFor:RowFormat
        //ExFor:Cell.CellFormat
        //ExFor:CellFormat
        //ExFor:CellFormat.Shading
        //ExSummary:Shows how to modify the format of rows and cells in a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("City");
        builder.insertCell();
        builder.write("Country");
        builder.endRow();
        builder.insertCell();
        builder.write("London");
        builder.insertCell();
        builder.write("U.K.");
        builder.endTable();

        // Use the first row's "RowFormat" property to modify the formatting
        // of the contents of all cells in this row.
        RowFormat rowFormat = table.getFirstRow().getRowFormat();
        rowFormat.setHeight(25.0);
        rowFormat.getBorders().getByBorderType(BorderType.BOTTOM).setColor(Color.RED);

        // Use the "CellFormat" property of the first cell in the last row to modify the formatting of that cell's contents.
        CellFormat cellFormat = table.getLastRow().getFirstCell().getCellFormat();
        cellFormat.setWidth(100.0);
        cellFormat.getShading().setBackgroundPatternColor(msColor.getOrange());

        doc.save(getArtifactsDir() + "Table.RowCellFormat.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.RowCellFormat.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals("City\u0007Country\u0007\u0007London\u0007U.K.\u0007\u0007", table.getText().trim());

        rowFormat = table.getFirstRow().getRowFormat();

        Assert.assertEquals(25.0d, rowFormat.getHeight());
        Assert.assertEquals(Color.RED.getRGB(), rowFormat.getBorders().getByBorderType(BorderType.BOTTOM).getColor().getRGB());

        cellFormat = table.getLastRow().getFirstCell().getCellFormat();

        Assert.assertEquals(110.8d, cellFormat.getWidth());
        Assert.assertEquals(msColor.getOrange().getRGB(), cellFormat.getShading().getBackgroundPatternColor().getRGB());
    }

    @Test
    public void displayContentOfTables() throws Exception
    {
        //ExStart
        //ExFor:Cell
        //ExFor:CellCollection
        //ExFor:CellCollection.Item(System.Int32)
        //ExFor:CellCollection.ToArray
        //ExFor:Row
        //ExFor:Row.Cells
        //ExFor:RowCollection
        //ExFor:RowCollection.Item(System.Int32)
        //ExFor:RowCollection.ToArray
        //ExFor:Table
        //ExFor:Table.Rows
        //ExFor:TableCollection.Item(System.Int32)
        //ExFor:TableCollection.ToArray
        //ExSummary:Shows how to iterate through all tables in the document and print the contents of each cell.
        Document doc = new Document(getMyDir() + "Tables.docx");
        TableCollection tables = doc.getFirstSection().getBody().getTables();

        Assert.assertEquals(2, tables.toArray().length);

        for (int i = 0; i < tables.getCount(); i++)
        {
            System.out.println("Start of Table {i}");

            RowCollection rows = tables.get(i).getRows();

            // We can use the "ToArray" method on a row collection to clone it into an array.
            Assert.assertEquals(rows, rows.toArray());
            Assert.assertNotSame(rows, rows.toArray());

            for (int j = 0; j < rows.getCount(); j++)
            {
                System.out.println("\tStart of Row {j}");

                CellCollection cells = rows.get(j).getCells();

                // We can use the "ToArray" method on a cell collection to clone it into an array.
                Assert.assertEquals(cells, cells.toArray());
                Assert.assertNotSame(cells, cells.toArray());

                for (int k = 0; k < cells.getCount(); k++)
                {
                    String cellText = cells.get(k).toString(SaveFormat.TEXT).trim();
                    System.out.println("\t\tContents of Cell:{k} = \"{cellText}\"");
                }

                System.out.println("\tEnd of Row {j}");
            }

            System.out.println("End of Table {i}\n");
        }
        //ExEnd
    }

    //ExStart
    //ExFor:Node.GetAncestor(NodeType)
    //ExFor:Node.GetAncestor(System.Type)
    //ExFor:Table.NodeType
    //ExFor:Cell.Tables
    //ExFor:TableCollection
    //ExFor:NodeCollection.Count
    //ExSummary:Shows how to find out if a tables are nested.
    @Test //ExSkip
    public void calculateDepthOfNestedTables() throws Exception
    {
        Document doc = new Document(getMyDir() + "Nested tables.docx");
        NodeCollection tables = doc.getChildNodes(NodeType.TABLE, true);
        Assert.assertEquals(5, tables.getCount()); //ExSkip

        for (int i = 0; i < tables.getCount(); i++)
        {
            Table table = (Table)tables.get(i);

            // Find out if any cells in the table have other tables as children.
            int count = getChildTableCount(table);
            System.out.println("Table #{0} has {1} tables directly within its cells",i,count);

            // Find out if the table is nested inside another table, and, if so, at what depth.
            int tableDepth = getNestedDepthOfTable(table);

            if (tableDepth > 0)
                System.out.println("Table #{0} is nested inside another table at depth of {1}",i,tableDepth);
            else
                System.out.println("Table #{0} is a non nested table (is not a child of another table)",i);
        }
    }

    /// <summary>
    /// Calculates what level a table is nested inside other tables.
    /// </summary>
    /// <returns>
    /// An integer indicating the nesting depth of the table (number of parent table nodes).
    /// </returns>
    private static int getNestedDepthOfTable(Table table)
    {
        int depth = 0;
        Node parent = table.getAncestor(table.getNodeType());

        while (parent != null)
        {
            depth++;
            parent = parent.getAncestor(Table.class);
        }

        return depth;
    }

    /// <summary>
    /// Determines if a table contains any immediate child table within its cells.
    /// Do not recursively traverse through those tables to check for further tables.
    /// </summary>
    /// <returns>
    /// Returns true if at least one child cell contains a table.
    /// Returns false if no cells in the table contain a table.
    /// </returns>
    private static int getChildTableCount(Table table)
    {
        int childTableCount = 0;

        for (Row row : table.getRows().<Row>OfType() !!Autoporter error: Undefined expression type )
        {
            for (Cell Cell : row.getCells().<Cell>OfType() !!Autoporter error: Undefined expression type )
            {
                TableCollection childTables = Cell.getTables();

                if (childTables.getCount() > 0)
                    childTableCount++;
            }
        }

        return childTableCount;
    }
    //ExEnd

    @Test
    public void ensureTableMinimum() throws Exception
    {
        //ExStart
        //ExFor:Table.EnsureMinimum
        //ExSummary:Shows how to ensure that a table node contains the nodes we need to add content.
        Document doc = new Document();
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);

        // Tables contain rows, which contain cells, which may contain paragraphs
        // with typical elements such as runs, shapes, and even other tables.
        // Our new table has none of these nodes, and we cannot add contents to it until it does.
        Assert.assertEquals(0, table.getChildNodes(NodeType.ANY, true).getCount());

        // Calling the "EnsureMinimum" method on a table will ensure that
        // the table has at least one row and one cell with an empty paragraph.
        table.ensureMinimum();
        table.getFirstRow().getFirstCell().getFirstParagraph().appendChild(new Run(doc, "Hello world!"));
        //ExEnd

        Assert.assertEquals(4, table.getChildNodes(NodeType.ANY, true).getCount());
    }

    @Test
    public void ensureRowMinimum() throws Exception
    {
        //ExStart
        //ExFor:Row.EnsureMinimum
        //ExSummary:Shows how to ensure a row node contains the nodes we need to begin adding content to it.
        Document doc = new Document();
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);
        Row row = new Row(doc);
        table.appendChild(row);

        // Rows contain cells, containing paragraphs with typical elements such as runs, shapes, and even other tables.
        // Our new row has none of these nodes, and we cannot add contents to it until it does.
        Assert.assertEquals(0, row.getChildNodes(NodeType.ANY, true).getCount());

        // Calling the "EnsureMinimum" method on a table will ensure that
        // the table has at least one cell with an empty paragraph.
        row.ensureMinimum();
        row.getFirstCell().getFirstParagraph().appendChild(new Run(doc, "Hello world!"));
        //ExEnd

        Assert.assertEquals(3, row.getChildNodes(NodeType.ANY, true).getCount());
    }

    @Test
    public void ensureCellMinimum() throws Exception
    {
        //ExStart
        //ExFor:Cell.EnsureMinimum
        //ExSummary:Shows how to ensure a cell node contains the nodes we need to begin adding content to it.
        Document doc = new Document();
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);
        Row row = new Row(doc);
        table.appendChild(row);
        Cell cell = new Cell(doc);
        row.appendChild(cell);

        // Cells may contain paragraphs with typical elements such as runs, shapes, and even other tables.
        // Our new cell does not have any paragraphs, and we cannot add contents such as run and shape nodes to it until it does.
        Assert.assertEquals(0, cell.getChildNodes(NodeType.ANY, true).getCount());

        // Calling the "EnsureMinimum" method on a cell will ensure that
        // the cell has at least one empty paragraph, which we can then add contents to.
        cell.ensureMinimum();
        cell.getFirstParagraph().appendChild(new Run(doc, "Hello world!"));
        //ExEnd

        Assert.assertEquals(2, cell.getChildNodes(NodeType.ANY, true).getCount());
    }

    @Test
    public void setOutlineBorders() throws Exception
    {
        //ExStart
        //ExFor:Table.Alignment
        //ExFor:TableAlignment
        //ExFor:Table.ClearBorders
        //ExFor:Table.ClearShading
        //ExFor:Table.SetBorder
        //ExFor:TextureIndex
        //ExFor:Table.SetShading
        //ExSummary:Shows how to apply an outline border to a table.
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        // Align the table to the center of the page.
        table.setAlignment(TableAlignment.CENTER);

        // Clear any existing borders and shading from the table.
        table.clearBorders();
        table.clearShading();

        // Add green borders to the outline of the table.
        table.setBorder(BorderType.LEFT, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.RIGHT, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.TOP, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.BOTTOM, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);

        // Fill the cells with a light green solid color.
        table.setShading(TextureIndex.TEXTURE_SOLID, msColor.getLightGreen(), msColor.Empty);

        doc.save(getArtifactsDir() + "Table.SetOutlineBorders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.SetOutlineBorders.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(TableAlignment.CENTER, table.getAlignment());

        BorderCollection borders = table.getFirstRow().getRowFormat().getBorders();

        Assert.assertEquals(msColor.getGreen().getRGB(), borders.getTop().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), borders.getLeft().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), borders.getRight().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), borders.getBottom().getColor().getRGB());
        Assert.assertNotEquals(msColor.getGreen().getRGB(), borders.getHorizontal().getColor().getRGB());
        Assert.assertNotEquals(msColor.getGreen().getRGB(), borders.getVertical().getColor().getRGB());
        Assert.assertEquals(msColor.getLightGreen().getRGB(), table.getFirstRow().getFirstCell().getCellFormat().getShading().getForegroundPatternColor().getRGB());
    }

    @Test
    public void setBorders() throws Exception
    {
        //ExStart
        //ExFor:Table.SetBorders
        //ExSummary:Shows how to format of all of a table's borders at once.
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        // Clear all existing borders from the table.
        table.clearBorders();

        // Set a single green line to serve as every outer and inner border of this table.
        table.setBorders(LineStyle.SINGLE, 1.5, msColor.getGreen());

        doc.save(getArtifactsDir() + "Table.SetBorders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.SetBorders.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);
        
        Assert.assertEquals(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getTop().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getLeft().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getRight().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getBottom().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getHorizontal().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getVertical().getColor().getRGB());
    }

    @Test
    public void rowFormat() throws Exception
    {
        //ExStart
        //ExFor:RowFormat
        //ExFor:Row.RowFormat
        //ExSummary:Shows how to modify formatting of a table row.
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        // Use the first row's "RowFormat" property to set formatting that modifies that entire row's appearance.
        Row firstRow = table.getFirstRow();
        firstRow.getRowFormat().getBorders().setLineStyle(LineStyle.NONE);
        firstRow.getRowFormat().setHeightRule(HeightRule.AUTO);
        firstRow.getRowFormat().setAllowBreakAcrossPages(true);

        doc.save(getArtifactsDir() + "Table.RowFormat.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.RowFormat.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(LineStyle.NONE, table.getFirstRow().getRowFormat().getBorders().getLineStyle());
        Assert.assertEquals(HeightRule.AUTO, table.getFirstRow().getRowFormat().getHeightRule());
        Assert.assertTrue(table.getFirstRow().getRowFormat().getAllowBreakAcrossPages());
    }

    @Test
    public void cellFormat() throws Exception
    {
        //ExStart
        //ExFor:CellFormat
        //ExFor:Cell.CellFormat
        //ExSummary:Shows how to modify formatting of a table cell.
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = doc.getFirstSection().getBody().getTables().get(0);
        Cell firstCell = table.getFirstRow().getFirstCell();

        // Use a cell's "CellFormat" property to set formatting that modifies the appearance of that cell.
        firstCell.getCellFormat().setWidth(30.0);
        firstCell.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        firstCell.getCellFormat().getShading().setForegroundPatternColor(msColor.getLightGreen());

        doc.save(getArtifactsDir() + "Table.CellFormat.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.CellFormat.docx");

        table = doc.getFirstSection().getBody().getTables().get(0);
        Assert.assertEquals(30, table.getFirstRow().getFirstCell().getCellFormat().getWidth());
        Assert.assertEquals(TextOrientation.DOWNWARD, table.getFirstRow().getFirstCell().getCellFormat().getOrientation());
        Assert.assertEquals(msColor.getLightGreen().getRGB(), table.getFirstRow().getFirstCell().getCellFormat().getShading().getForegroundPatternColor().getRGB());
    }

    @Test
    public void getDistance() throws Exception
    {
        //ExStart
        //ExFor:Table.DistanceBottom
        //ExFor:Table.DistanceLeft
        //ExFor:Table.DistanceRight
        //ExFor:Table.DistanceTop
        //ExSummary:Shows the minimum distance operations between table boundaries and text.
        Document doc = new Document(getMyDir() + "Table wrapped by text.docx");

        Table table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(25.9d, table.getDistanceTop());
        Assert.assertEquals(25.9d, table.getDistanceBottom());
        Assert.assertEquals(17.3d, table.getDistanceLeft());
        Assert.assertEquals(17.3d, table.getDistanceRight());
        //ExEnd
    }

    @Test
    public void borders() throws Exception
    {
        //ExStart
        //ExFor:Table.ClearBorders
        //ExSummary:Shows how to remove all borders from a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Hello world!");
        builder.endTable();

        // Modify the color and thickness of the top border.
        Border topBorder = table.getFirstRow().getRowFormat().getBorders().getByBorderType(BorderType.TOP);
        table.setBorder(BorderType.TOP, LineStyle.DOUBLE, 1.5, Color.RED, true);

        Assert.assertEquals(1.5d, topBorder.getLineWidth());
        Assert.assertEquals(Color.RED.getRGB(), topBorder.getColor().getRGB());
        Assert.assertEquals(LineStyle.DOUBLE, topBorder.getLineStyle());

        // Clear the borders of all cells in the table, and then save the document.
        table.clearBorders();
        Assert.<AssertionError>Throws(() => Assert.assertEquals(msColor.Empty.getRGB(), topBorder.getColor().getRGB())); //ExSkip
        doc.save(getArtifactsDir() + "Table.ClearBorders.docx");

        // Verify the values of the table's properties after re-opening the document.
        doc = new Document(getArtifactsDir() + "Table.ClearBorders.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);
        topBorder = table.getFirstRow().getRowFormat().getBorders().getByBorderType(BorderType.TOP);

        Assert.assertEquals(0.0d, topBorder.getLineWidth());
        Assert.assertEquals(msColor.Empty.getRGB(), topBorder.getColor().getRGB());
        Assert.assertEquals(LineStyle.NONE, topBorder.getLineStyle());
        //ExEnd
    }

    @Test
    public void replaceCellText() throws Exception
    {
        //ExStart
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace all instances of String of text in a table and cell.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Carrots");
        builder.insertCell();
        builder.write("50");
        builder.endRow();
        builder.insertCell();
        builder.write("Potatoes");
        builder.insertCell();
        builder.write("50");
        builder.endTable();

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(true);
        options.setFindWholeWordsOnly(true);

        // Perform a find-and-replace operation on an entire table.
        table.getRange().replace("Carrots", "Eggs", options);

        // Perform a find-and-replace operation on the last cell of the last row of the table.
        table.getLastRow().getLastCell().getRange().replace("50", "20", options);

        Assert.assertEquals("Eggs\u000750\u0007\u0007" +
                        "Potatoes\u000720\u0007\u0007", table.getText().trim());
        //ExEnd
    }

    @Test (dataProvider = "removeParagraphTextAndMarkDataProvider")
    public void removeParagraphTextAndMark(boolean isSmartParagraphBreakReplacement) throws Exception
    {
        //ExStart
        //ExFor:FindReplaceOptions.SmartParagraphBreakReplacement
        //ExSummary:Shows how to remove paragraph from a table cell with a nested table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create table with paragraph and inner table in first cell.
        builder.startTable();
        builder.insertCell();
        builder.write("TEXT1");
        builder.startTable();
        builder.insertCell();
        builder.endTable();
        builder.endTable();
        builder.writeln();

        FindReplaceOptions options = new FindReplaceOptions();
        // When the following option is set to 'true', Aspose.Words will remove paragraph's text
        // completely with its paragraph mark. Otherwise, Aspose.Words will mimic Word and remove
        // only paragraph's text and leaves the paragraph mark intact (when a table follows the text).
        options.setSmartParagraphBreakReplacement(isSmartParagraphBreakReplacement);
        doc.getRange().replaceInternal(new Regex("TEXT1&p"), "", options);

        doc.save(getArtifactsDir() + "Table.RemoveParagraphTextAndMark.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.RemoveParagraphTextAndMark.docx");

        Assert.assertEquals(isSmartParagraphBreakReplacement ? 1 : 2,
            doc.getFirstSection().getBody().getTables().get(0).getRows().get(0).getCells().get(0).getParagraphs().getCount());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "removeParagraphTextAndMarkDataProvider")
	public static Object[][] removeParagraphTextAndMarkDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void printTableRange() throws Exception
    {
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = doc.getFirstSection().getBody().getTables().get(0);

        // The range text will include control characters such as "\a" for a cell.
        // You can call ToString on the desired node to retrieve the plain text content.

        // Print the plain text range of the table to the screen.
        System.out.println("Contents of the table: ");
        System.out.println(table.getRange().getText());
        
        // Print the contents of the second row to the screen.
        System.out.println("\nContents of the row: ");
        System.out.println(table.getRows().get(1).getRange().getText());

        // Print the contents of the last cell in the table to the screen.
        System.out.println("\nContents of the cell: ");
        System.out.println(table.getLastRow().getLastCell().getRange().getText());
        
        Assert.assertEquals("\u0007Column 1\u0007Column 2\u0007Column 3\u0007Column 4\u0007\u0007", table.getRows().get(1).getRange().getText());
        Assert.assertEquals("Cell 12 contents\u0007", table.getLastRow().getLastCell().getRange().getText());
    }

    @Test
    public void cloneTable() throws Exception
    {
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = doc.getFirstSection().getBody().getTables().get(0);

        Table tableClone = (Table) table.deepClone(true);

        // Insert the cloned table into the document after the original.
        table.getParentNode().insertAfter(tableClone, table);

        // Insert an empty paragraph between the two tables.
        table.getParentNode().insertAfter(new Paragraph(doc), table);

        doc.save(getArtifactsDir() + "Table.CloneTable.doc");
        
        Assert.assertEquals(3, doc.getChildNodes(NodeType.TABLE, true).getCount());
        Assert.assertEquals(table.getRange().getText(), tableClone.getRange().getText());

        for (Cell cell : tableClone.getChildNodes(NodeType.CELL, true).<Cell>OfType() !!Autoporter error: Undefined expression type )
            cell.removeAllChildren();
        
        Assert.assertEquals("", tableClone.toString(SaveFormat.TEXT).trim());
    }

    @Test (dataProvider = "allowBreakAcrossPagesDataProvider")
    public void allowBreakAcrossPages(boolean allowBreakAcrossPages) throws Exception
    {
        //ExStart
        //ExFor:RowFormat.AllowBreakAcrossPages
        //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
        Document doc = new Document(getMyDir() + "Table spanning two pages.docx");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        // Set the "AllowBreakAcrossPages" property to "false" to keep the row
        // in one piece if a table spans two pages, which break up along that row.
        // If the row is too big to fit in one page, Microsoft Word will push it down to the next page.
        // Set the "AllowBreakAcrossPages" property to "true" to allow the row to break up across two pages.
        for (Row row : table.<Row>OfType() !!Autoporter error: Undefined expression type )
            row.getRowFormat().setAllowBreakAcrossPages(allowBreakAcrossPages);

        doc.save(getArtifactsDir() + "Table.AllowBreakAcrossPages.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.AllowBreakAcrossPages.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.AreEqual(3, table.getRows().Count(r => ((Row)r).RowFormat.AllowBreakAcrossPages == allowBreakAcrossPages));
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "allowBreakAcrossPagesDataProvider")
	public static Object[][] allowBreakAcrossPagesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "allowAutoFitOnTableDataProvider")
    public void allowAutoFitOnTable(boolean allowAutoFit) throws Exception
    {
        //ExStart
        //ExFor:Table.AllowAutoFit
        //ExSummary:Shows how to enable/disable automatic table cell resizing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPoints(100.0));
        builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                      "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.AUTO);
        builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                      "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder.endRow();
        builder.endTable();

        // Set the "AllowAutoFit" property to "false" to get the table to maintain the dimensions
        // of all its rows and cells, and truncate contents if they get too large to fit.
        // Set the "AllowAutoFit" property to "true" to allow the table to change its cells' width and height
        // to accommodate their contents.
        table.setAllowAutoFit(allowAutoFit);

        doc.save(getArtifactsDir() + "Table.AllowAutoFitOnTable.html");
        //ExEnd

        if (allowAutoFit)
        {
            TestUtil.fileContainsString(
                "<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">",
                getArtifactsDir() + "Table.AllowAutoFitOnTable.html");
            TestUtil.fileContainsString(
                "<td style=\"border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">",
                getArtifactsDir() + "Table.AllowAutoFitOnTable.html");
        }
        else
        {
            TestUtil.fileContainsString(
                "<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">",
                getArtifactsDir() + "Table.AllowAutoFitOnTable.html");
            TestUtil.fileContainsString(
                "<td style=\"width:7.2pt; border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">",
                getArtifactsDir() + "Table.AllowAutoFitOnTable.html");
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "allowAutoFitOnTableDataProvider")
	public static Object[][] allowAutoFitOnTableDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void keepTableTogether() throws Exception
    {
        //ExStart
        //ExFor:ParagraphFormat.KeepWithNext
        //ExFor:Row.IsLastRow
        //ExFor:Paragraph.IsEndOfCell
        //ExFor:Paragraph.IsInCell
        //ExFor:Cell.ParentRow
        //ExFor:Cell.Paragraphs
        //ExSummary:Shows how to set a table to stay together on the same page.
        Document doc = new Document(getMyDir() + "Table spanning two pages.docx");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        // Enabling KeepWithNext for every paragraph in the table except for the
        // last ones in the last row will prevent the table from splitting across multiple pages.
        for (Cell cell : table.getChildNodes(NodeType.CELL, true).<Cell>OfType() !!Autoporter error: Undefined expression type )
            for (Paragraph para : cell.getParagraphs().<Paragraph>OfType() !!Autoporter error: Undefined expression type )
            {
                Assert.assertTrue(para.isInCell());

                if (!(cell.getParentRow().isLastRow() && para.isEndOfCell()))
                    para.getParagraphFormat().setKeepWithNext(true);
            }

        doc.save(getArtifactsDir() + "Table.KeepTableTogether.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.KeepTableTogether.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        for (Paragraph para : table.getChildNodes(NodeType.PARAGRAPH, true).<Paragraph>OfType() !!Autoporter error: Undefined expression type )
            if (para.isEndOfCell() && ((Cell)para.getParentNode()).getParentRow().isLastRow())
                Assert.assertFalse(para.getParagraphFormat().getKeepWithNext());
            else
                Assert.assertTrue(para.getParagraphFormat().getKeepWithNext());
    }

    @Test
    public void getIndexOfTableElements() throws Exception
    {
        //ExStart
        //ExFor:NodeCollection.IndexOf(Node)
        //ExSummary:Shows how to get the index of a node in a collection.
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = doc.getFirstSection().getBody().getTables().get(0);
        NodeCollection allTables = doc.getChildNodes(NodeType.TABLE, true);

        Assert.assertEquals(0, allTables.indexOf(table));

        Row row = table.getRows().get(2);

        Assert.assertEquals(2, table.indexOf(row));

        Cell cell = row.getLastCell();

        Assert.assertEquals(4, row.indexOf(cell));
        //ExEnd
    }

    @Test
    public void getPreferredWidthTypeAndValue() throws Exception
    {
        //ExStart
        //ExFor:PreferredWidthType
        //ExFor:PreferredWidth.Type
        //ExFor:PreferredWidth.Value
        //ExSummary:Shows how to verify the preferred width type and value of a table cell.
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = doc.getFirstSection().getBody().getTables().get(0);
        Cell firstCell = table.getFirstRow().getFirstCell();

        Assert.assertEquals(PreferredWidthType.PERCENT, firstCell.getCellFormat().getPreferredWidth().getType());
        Assert.assertEquals(11.16d, firstCell.getCellFormat().getPreferredWidth().getValue());
        //ExEnd
    }

    @Test (dataProvider = "allowCellSpacingDataProvider")
    public void allowCellSpacing(boolean allowCellSpacing) throws Exception
    {
        //ExStart
        //ExFor:Table.AllowCellSpacing
        //ExFor:Table.CellSpacing
        //ExSummary:Shows how to enable spacing between individual cells in a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Animal");
        builder.insertCell();
        builder.write("Class");
        builder.endRow();
        builder.insertCell();
        builder.write("Dog");
        builder.insertCell();
        builder.write("Mammal");
        builder.endTable();

        table.setCellSpacing(3.0);

        // Set the "AllowCellSpacing" property to "true" to enable spacing between cells
        // with a magnitude equal to the value of the "CellSpacing" property, in points.
        // Set the "AllowCellSpacing" property to "false" to disable cell spacing
        // and ignore the value of the "CellSpacing" property.
        table.setAllowCellSpacing(allowCellSpacing);

        doc.save(getArtifactsDir() + "Table.AllowCellSpacing.html");

        // Adjusting the "CellSpacing" property will automatically enable cell spacing.
        table.setCellSpacing(5.0);

        Assert.assertTrue(table.getAllowCellSpacing());
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.AllowCellSpacing.html");
        table = (Table)doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(allowCellSpacing, table.getAllowCellSpacing());

        if (allowCellSpacing)
            Assert.assertEquals(3.0d, table.getCellSpacing());
        else
            Assert.assertEquals(0.0d, table.getCellSpacing());

        TestUtil.fileContainsString(
            allowCellSpacing
                ? "<td style=\"border-style:solid; border-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border:0.5pt single\">"
                : "<td style=\"border-right-style:solid; border-right-width:0.75pt; border-bottom-style:solid; border-bottom-width:0.75pt; " +
                  "padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-bottom:0.5pt single; -aw-border-right:0.5pt single\">",
            getArtifactsDir() + "Table.AllowCellSpacing.html");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "allowCellSpacingDataProvider")
	public static Object[][] allowCellSpacingDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    //ExStart
    //ExFor:Table
    //ExFor:Row
    //ExFor:Cell
    //ExFor:Table.#ctor(DocumentBase)
    //ExFor:Table.Title
    //ExFor:Table.Description
    //ExFor:Row.#ctor(DocumentBase)
    //ExFor:Cell.#ctor(DocumentBase)
    //ExFor:Cell.FirstParagraph
    //ExSummary:Shows how to build a nested table without using a document builder.
    @Test //ExSkip
    public void createNestedTable() throws Exception
    {
        Document doc = new Document();

        // Create the outer table with three rows and four columns, and then add it to the document.
        Table outerTable = createTable(doc, 3, 4, "Outer Table");
        doc.getFirstSection().getBody().appendChild(outerTable);

        // Create another table with two rows and two columns and then insert it into the first table's first cell.
        Table innerTable = createTable(doc, 2, 2, "Inner Table");
        outerTable.getFirstRow().getFirstCell().appendChild(innerTable);

        doc.save(getArtifactsDir() + "Table.CreateNestedTable.docx");
        testCreateNestedTable(new Document(getArtifactsDir() + "Table.CreateNestedTable.docx")); //ExSkip
    }

    /// <summary>
    /// Creates a new table in the document with the given dimensions and text in each cell.
    /// </summary>
    private static Table createTable(Document doc, int rowCount, int cellCount, String cellText) throws Exception
    {
        Table table = new Table(doc);

        for (int rowId = 1; rowId <= rowCount; rowId++)
        {
            Row row = new Row(doc);
            table.appendChild(row);

            for (int cellId = 1; cellId <= cellCount; cellId++)
            {
                Cell cell = new Cell(doc);
                cell.appendChild(new Paragraph(doc));
                cell.getFirstParagraph().appendChild(new Run(doc, cellText));

                row.appendChild(cell);
            }
        }

        // You can use the "Title" and "Description" properties to add a title and description respectively to your table.
        // The table must have at least one row before we can use these properties.
        // These properties are meaningful for ISO / IEC 29500 compliant .docx documents (see the OoxmlCompliance class).
        // If we save the document to pre-ISO/IEC 29500 formats, Microsoft Word ignores these properties.
        table.setTitle("Aspose table title");
        table.setDescription("Aspose table description");

        return table;
    }
    //ExEnd

    private void testCreateNestedTable(Document doc)
    {
        Table outerTable = doc.getFirstSection().getBody().getTables().get(0);
        Table innerTable = (Table)doc.getChild(NodeType.TABLE, 1, true);

        Assert.assertEquals(2, doc.getChildNodes(NodeType.TABLE, true).getCount());
        Assert.assertEquals(1, outerTable.getFirstRow().getFirstCell().getTables().getCount());
        Assert.assertEquals(16, outerTable.getChildNodes(NodeType.CELL, true).getCount());
        Assert.assertEquals(4, innerTable.getChildNodes(NodeType.CELL, true).getCount());
        Assert.assertEquals("Aspose table title", innerTable.getTitle());
        Assert.assertEquals("Aspose table description", innerTable.getDescription());
    }

    //ExStart
    //ExFor:CellFormat.HorizontalMerge
    //ExFor:CellFormat.VerticalMerge
    //ExFor:CellMerge
    //ExSummary:Prints the horizontal and vertical merge type of a cell.
    @Test //ExSkip
    public void checkCellsMerged() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table with merged cells.docx");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        for (Row row : table.getRows().<Row>OfType() !!Autoporter error: Undefined expression type )
            for (Cell cell : row.getCells().<Cell>OfType() !!Autoporter error: Undefined expression type )
                System.out.println(printCellMergeType(cell));
        Assert.assertEquals("The cell at R1, C1 is vertically merged", printCellMergeType(table.getFirstRow().getFirstCell())); //ExSkip
    }

    @Test (enabled = false)
    public String printCellMergeType(Cell cell)
    {
        boolean isHorizontallyMerged = cell.getCellFormat().getHorizontalMerge() != CellMerge.NONE;
        boolean isVerticallyMerged = cell.getCellFormat().getVerticalMerge() != CellMerge.NONE;
        String cellLocation =
            $"R{cell.ParentRow.ParentTable.IndexOf(cell.ParentRow) + 1}, C{cell.ParentRow.IndexOf(cell) + 1}";

        if (isHorizontallyMerged && isVerticallyMerged)
            return $"The cell at {cellLocation} is both horizontally and vertically merged";
        if (isHorizontallyMerged)
            return $"The cell at {cellLocation} is horizontally merged.";

        return isVerticallyMerged ? $"The cell at {cellLocation} is vertically merged" : $"The cell at {cellLocation} is not merged";
    }
    //ExEnd

    @Test
    public void mergeCellRange() throws Exception
    {
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = doc.getFirstSection().getBody().getTables().get(0);

        // We want to merge the range of cells found in between these two cells.
        Cell cellStartRange = table.getRows().get(2).getCells().get(2);
        Cell cellEndRange = table.getRows().get(3).getCells().get(3);

        // Merge all the cells between the two specified cells into one.
        mergeCells(cellStartRange, cellEndRange);

        doc.save(getArtifactsDir() + "Table.MergeCellRange.doc");

        int mergedCellsCount = 0;
        for (Node node : (Iterable<Node>) table.getChildNodes(NodeType.CELL, true))
        {
            Cell cell = (Cell) node;
            if (cell.getCellFormat().getHorizontalMerge() != CellMerge.NONE ||
                cell.getCellFormat().getVerticalMerge() != CellMerge.NONE)
                mergedCellsCount++;
        }

        Assert.assertEquals(4, mergedCellsCount);
        Assert.assertTrue(table.getRows().get(2).getCells().get(2).getCellFormat().getHorizontalMerge() == CellMerge.FIRST);
        Assert.assertTrue(table.getRows().get(2).getCells().get(2).getCellFormat().getVerticalMerge() == CellMerge.FIRST);
        Assert.assertTrue(table.getRows().get(3).getCells().get(3).getCellFormat().getHorizontalMerge() == CellMerge.PREVIOUS);
        Assert.assertTrue(table.getRows().get(3).getCells().get(3).getCellFormat().getVerticalMerge() == CellMerge.PREVIOUS);
    }

    /// <summary>
    /// Merges the range of cells found between the two specified cells both horizontally and vertically.
    /// Can span over multiple rows.
    /// </summary>
    @Test (enabled = false)
    public static void mergeCells(Cell startCell, Cell endCell)
    {
        Table parentTable = startCell.getParentRow().getParentTable();

        // Find the row and cell indices for the start and end cells.
        /*Point*/long startCellPos = msPoint.ctor(startCell.getParentRow().indexOf(startCell),
            parentTable.indexOf(startCell.getParentRow()));
        /*Point*/long endCellPos = msPoint.ctor(endCell.getParentRow().indexOf(endCell), parentTable.indexOf(endCell.getParentRow()));

        // Create a range of cells to be merged based on these indices.
        // Inverse each index if the end cell is before the start cell.
        Rectangle mergeRange = new Rectangle(
            Math.min(msPoint.getX(startCellPos), msPoint.getX(endCellPos)),
            Math.min(msPoint.getY(startCellPos), msPoint.getY(endCellPos)),
            Math.abs(msPoint.getX(endCellPos) - msPoint.getX(startCellPos)) + 1,
            Math.abs(msPoint.getY(endCellPos) - msPoint.getY(startCellPos)) + 1);

        for (Row row : parentTable.getRows().<Row>OfType() !!Autoporter error: Undefined expression type )
        {
            for (Cell cell : row.getCells().<Cell>OfType() !!Autoporter error: Undefined expression type )
            {
                /*Point*/long currentPos = msPoint.ctor(row.indexOf(cell), parentTable.indexOf(row));

                // Check if the current cell is inside our merge range, then merge it.
                if (mergeRange.contains(currentPos))
                {
                    cell.getCellFormat().setHorizontalMerge(msPoint.getX(currentPos) == mergeRange.getX() ? CellMerge.FIRST : CellMerge.PREVIOUS);
                    cell.getCellFormat().setVerticalMerge(msPoint.getY(currentPos) == mergeRange.getY() ? CellMerge.FIRST : CellMerge.PREVIOUS);
                }
            }
        }
    }

    @Test
    public void combineTables() throws Exception
    {
        //ExStart
        //ExFor:Cell.CellFormat
        //ExFor:CellFormat.Borders
        //ExFor:Table.Rows
        //ExFor:Table.FirstRow
        //ExFor:CellFormat.ClearFormatting
        //ExFor:CompositeNode.HasChildNodes
        //ExSummary:Shows how to combine the rows from two tables into one.
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Below are two ways of getting a table from a document.
        // 1 -  From the "Tables" collection of a Body node:
        Table firstTable = doc.getFirstSection().getBody().getTables().get(0);

        // 2 -  Using the "GetChild" method:
        Table secondTable = (Table)doc.getChild(NodeType.TABLE, 1, true);

        // Append all rows from the current table to the next.
        while (secondTable.hasChildNodes())
            firstTable.getRows().add(secondTable.getFirstRow());

        // Remove the empty table container.
        secondTable.remove();

        doc.save(getArtifactsDir() + "Table.CombineTables.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.CombineTables.docx");

        Assert.assertEquals(1, doc.getChildNodes(NodeType.TABLE, true).getCount());
        Assert.assertEquals(9, doc.getFirstSection().getBody().getTables().get(0).getRows().getCount());
        Assert.assertEquals(42, doc.getFirstSection().getBody().getTables().get(0).getChildNodes(NodeType.CELL, true).getCount());
    }

    @Test
    public void splitTable() throws Exception
    {
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table firstTable = doc.getFirstSection().getBody().getTables().get(0);

        // We will split the table at the third row (inclusive).
        Row row = firstTable.getRows().get(2);

        // Create a new container for the split table.
        Table table = (Table) firstTable.deepClone(false);

        // Insert the container after the original.
        firstTable.getParentNode().insertAfter(table, firstTable);

        // Add a buffer paragraph to ensure the tables stay apart.
        firstTable.getParentNode().insertAfter(new Paragraph(doc), firstTable);

        Row currentRow;

        do
        {
            currentRow = firstTable.getLastRow();
            table.prependChild(currentRow);
        } while (currentRow != row);

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(row, table.getFirstRow());
        Assert.assertEquals(2, firstTable.getRows().getCount());
        Assert.assertEquals(3, table.getRows().getCount());
        Assert.assertEquals(3, doc.getChildNodes(NodeType.TABLE, true).getCount());
    }

    @Test
    public void wrapText() throws Exception
    {
        //ExStart
        //ExFor:Table.TextWrapping
        //ExFor:TextWrapping
        //ExSummary:Shows how to work with table text wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.endTable();
        table.setPreferredWidth(PreferredWidth.fromPoints(300.0));

        builder.getFont().setSize(16.0);
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Set the "TextWrapping" property to "TextWrapping.Around" to get the table to wrap text around it,
        // and push it down into the paragraph below by setting the position.
        table.setTextWrapping(TextWrapping.AROUND);
        table.setAbsoluteHorizontalDistance(100.0);
        table.setAbsoluteVerticalDistance(20.0);

        doc.save(getArtifactsDir() + "Table.WrapText.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.WrapText.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(TextWrapping.AROUND, table.getTextWrapping());
        Assert.assertEquals(100.0d, table.getAbsoluteHorizontalDistance());
        Assert.assertEquals(20.0d, table.getAbsoluteVerticalDistance());
    }

    @Test
    public void getFloatingTableProperties() throws Exception
    {
        //ExStart
        //ExFor:Table.HorizontalAnchor
        //ExFor:Table.VerticalAnchor
        //ExFor:Table.AllowOverlap
        //ExFor:ShapeBase.AllowOverlap
        //ExSummary:Shows how to work with floating tables properties.
        Document doc = new Document(getMyDir() + "Table wrapped by text.docx");

        Table table = doc.getFirstSection().getBody().getTables().get(0);

        if (table.getTextWrapping() == TextWrapping.AROUND)
        {
            Assert.assertEquals(RelativeHorizontalPosition.MARGIN, table.getHorizontalAnchor());
            Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, table.getVerticalAnchor());
            Assert.assertEquals(false, table.getAllowOverlap());

            // Only Margin, Page, Column available in RelativeHorizontalPosition for HorizontalAnchor setter.
            // The ArgumentException will be thrown for any other values.
            table.setHorizontalAnchor(RelativeHorizontalPosition.COLUMN);

            // Only Margin, Page, Paragraph available in RelativeVerticalPosition for VerticalAnchor setter.
            // The ArgumentException will be thrown for any other values.
            table.setVerticalAnchor(RelativeVerticalPosition.PAGE);
        }
        //ExEnd
    }

    @Test
    public void changeFloatingTableProperties() throws Exception
    {
        //ExStart
        //ExFor:Table.RelativeHorizontalAlignment
        //ExFor:Table.RelativeVerticalAlignment
        //ExFor:Table.AbsoluteHorizontalDistance
        //ExFor:Table.AbsoluteVerticalDistance
        //ExSummary:Shows how set the location of floating tables.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Table 1, cell 1");
        builder.endTable();
        table.setPreferredWidth(PreferredWidth.fromPoints(300.0));

        // Set the table's location to a place on the page, such as, in this case, the bottom right corner.
        table.setRelativeVerticalAlignment(VerticalAlignment.BOTTOM);
        table.setRelativeHorizontalAlignment(HorizontalAlignment.RIGHT);

        table = builder.startTable();
        builder.insertCell();
        builder.write("Table 2, cell 1");
        builder.endTable();
        table.setPreferredWidth(PreferredWidth.fromPoints(300.0));

        // We can also set a horizontal and vertical offset in points from the paragraph's location where we inserted the table. 
        table.setAbsoluteVerticalDistance(50.0);
        table.setAbsoluteHorizontalDistance(100.0);

        doc.save(getArtifactsDir() + "Table.ChangeFloatingTableProperties.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.ChangeFloatingTableProperties.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(VerticalAlignment.BOTTOM, table.getRelativeVerticalAlignment());
        Assert.assertEquals(HorizontalAlignment.RIGHT, table.getRelativeHorizontalAlignment());

        table = (Table)doc.getChild(NodeType.TABLE, 1, true);

        Assert.assertEquals(50.0d, table.getAbsoluteVerticalDistance());
        Assert.assertEquals(100.0d, table.getAbsoluteHorizontalDistance());
    }

    @Test
    public void tableStyleCreation() throws Exception
    {
        //ExStart
        //ExFor:Table.Bidi
        //ExFor:Table.CellSpacing
        //ExFor:Table.Style
        //ExFor:Table.StyleName
        //ExFor:TableStyle
        //ExFor:TableStyle.AllowBreakAcrossPages
        //ExFor:TableStyle.Bidi
        //ExFor:TableStyle.CellSpacing
        //ExFor:TableStyle.BottomPadding
        //ExFor:TableStyle.LeftPadding
        //ExFor:TableStyle.RightPadding
        //ExFor:TableStyle.TopPadding
        //ExFor:TableStyle.Shading
        //ExFor:TableStyle.Borders
        //ExFor:TableStyle.VerticalAlignment
        //ExSummary:Shows how to create custom style settings for the table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
 
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Name");
        builder.insertCell();
        builder.write("مرحبًا");
        builder.endRow();
        builder.insertCell();
        builder.insertCell();
        builder.endTable();
 
        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.setAllowBreakAcrossPages(true);
        tableStyle.setBidi(true);
        tableStyle.setCellSpacing(5.0);
        tableStyle.setBottomPadding(20.0);
        tableStyle.setLeftPadding(5.0);
        tableStyle.setRightPadding(10.0);
        tableStyle.setTopPadding(20.0);
        tableStyle.getShading().setBackgroundPatternColor(Color.AntiqueWhite);
        tableStyle.getBorders().setColor(Color.BLUE);
        tableStyle.getBorders().setLineStyle(LineStyle.DOT_DASH);
        tableStyle.setVerticalAlignment(CellVerticalAlignment.CENTER);

        table.setStyle(tableStyle);

        // Setting the style properties of a table may affect the properties of the table itself.
        Assert.assertTrue(table.getBidi());
        Assert.assertEquals(5.0d, table.getCellSpacing());
        Assert.assertEquals("MyTableStyle1", table.getStyleName());

        doc.save(getArtifactsDir() + "Table.TableStyleCreation.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.TableStyleCreation.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertTrue(table.getBidi());
        Assert.assertEquals(5.0d, table.getCellSpacing());
        Assert.assertEquals("MyTableStyle1", table.getStyleName());
        Assert.assertEquals(20.0d, tableStyle.getBottomPadding());
        Assert.assertEquals(5.0d, tableStyle.getLeftPadding());
        Assert.assertEquals(10.0d, tableStyle.getRightPadding());
        Assert.assertEquals(20.0d, tableStyle.getTopPadding());
        Assert.AreEqual(6, table.getFirstRow().getRowFormat().getBorders().Count(b => b.Color.ToArgb() == Color.Blue.ToArgb()));
        Assert.assertEquals(CellVerticalAlignment.CENTER, tableStyle.getVerticalAlignment());

        tableStyle = (TableStyle)doc.getStyles().get("MyTableStyle1");

        Assert.assertTrue(tableStyle.getAllowBreakAcrossPages());
        Assert.assertTrue(tableStyle.getBidi());
        Assert.assertEquals(5.0d, tableStyle.getCellSpacing());
        Assert.assertEquals(20.0d, tableStyle.getBottomPadding());
        Assert.assertEquals(5.0d, tableStyle.getLeftPadding());
        Assert.assertEquals(10.0d, tableStyle.getRightPadding());
        Assert.assertEquals(20.0d, tableStyle.getTopPadding());
        Assert.assertEquals(Color.AntiqueWhite.getRGB(), tableStyle.getShading().getBackgroundPatternColor().getRGB());
        Assert.assertEquals(Color.BLUE.getRGB(), tableStyle.getBorders().getColor().getRGB());
        Assert.assertEquals(LineStyle.DOT_DASH, tableStyle.getBorders().getLineStyle());
        Assert.assertEquals(CellVerticalAlignment.CENTER, tableStyle.getVerticalAlignment());
    }

    @Test
    public void setTableAlignment() throws Exception
    {
        //ExStart
        //ExFor:TableStyle.Alignment
        //ExFor:TableStyle.LeftIndent
        //ExSummary:Shows how to set the position of a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two ways of aligning a table horizontally.
        // 1 -  Use the "Alignment" property to align it to a location on the page, such as the center:
        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.setAlignment(TableAlignment.CENTER);
        tableStyle.getBorders().setColor(Color.BLUE);
        tableStyle.getBorders().setLineStyle(LineStyle.SINGLE);

        // Insert a table and apply the style we created to it.
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Aligned to the center of the page");
        builder.endTable();
        table.setPreferredWidth(PreferredWidth.fromPoints(300.0));
        
        table.setStyle(tableStyle);

        // 2 -  Use the "LeftIndent" to specify an indent from the left margin of the page:
        tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle2");
        tableStyle.setLeftIndent(55.0);
        tableStyle.getBorders().setColor(msColor.getGreen());
        tableStyle.getBorders().setLineStyle(LineStyle.SINGLE);

        table = builder.startTable();
        builder.insertCell();
        builder.write("Aligned according to left indent");
        builder.endTable();
        table.setPreferredWidth(PreferredWidth.fromPoints(300.0));

        table.setStyle(tableStyle);

        doc.save(getArtifactsDir() + "Table.SetTableAlignment.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.SetTableAlignment.docx");

        tableStyle = (TableStyle)doc.getStyles().get("MyTableStyle1");

        Assert.assertEquals(TableAlignment.CENTER, tableStyle.getAlignment());
        Assert.assertEquals(tableStyle, doc.getFirstSection().getBody().getTables().get(0).getStyle());

        tableStyle = (TableStyle)doc.getStyles().get("MyTableStyle2");

        Assert.assertEquals(55.0d, tableStyle.getLeftIndent());
        Assert.assertEquals(tableStyle, ((Table)doc.getChild(NodeType.TABLE, 1, true)).getStyle());
    }

    @Test
    public void conditionalStyles() throws Exception
    {
        //ExStart
        //ExFor:ConditionalStyle
        //ExFor:ConditionalStyle.Shading
        //ExFor:ConditionalStyle.Borders
        //ExFor:ConditionalStyle.ParagraphFormat
        //ExFor:ConditionalStyle.BottomPadding
        //ExFor:ConditionalStyle.LeftPadding
        //ExFor:ConditionalStyle.RightPadding
        //ExFor:ConditionalStyle.TopPadding
        //ExFor:ConditionalStyle.Font
        //ExFor:ConditionalStyle.Type
        //ExFor:ConditionalStyleCollection.GetEnumerator
        //ExFor:ConditionalStyleCollection.FirstRow
        //ExFor:ConditionalStyleCollection.LastRow
        //ExFor:ConditionalStyleCollection.LastColumn
        //ExFor:ConditionalStyleCollection.Count
        //ExFor:ConditionalStyleCollection
        //ExFor:ConditionalStyleCollection.BottomLeftCell
        //ExFor:ConditionalStyleCollection.BottomRightCell
        //ExFor:ConditionalStyleCollection.EvenColumnBanding
        //ExFor:ConditionalStyleCollection.EvenRowBanding
        //ExFor:ConditionalStyleCollection.FirstColumn
        //ExFor:ConditionalStyleCollection.Item(ConditionalStyleType)
        //ExFor:ConditionalStyleCollection.Item(Int32)
        //ExFor:ConditionalStyleCollection.OddColumnBanding
        //ExFor:ConditionalStyleCollection.OddRowBanding
        //ExFor:ConditionalStyleCollection.TopLeftCell
        //ExFor:ConditionalStyleCollection.TopRightCell
        //ExFor:ConditionalStyleType
        //ExFor:TableStyle.ConditionalStyles
        //ExSummary:Shows how to work with certain area styles of a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.endRow();
        builder.insertCell();
        builder.write("Cell 3");
        builder.insertCell();
        builder.write("Cell 4");
        builder.endTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");

        // Conditional styles are formatting changes that affect only some of the table's cells
        // based on a predicate, such as the cells being in the last row.
        // Below are three ways of accessing a table style's conditional styles from the "ConditionalStyles" collection.
        // 1 -  By style type:
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.FIRST_ROW).getShading().setBackgroundPatternColor(msColor.getAliceBlue());

        // 2 -  By index:
        tableStyle.getConditionalStyles().get(0).getBorders().setColor(Color.BLACK);
        tableStyle.getConditionalStyles().get(0).getBorders().setLineStyle(LineStyle.DOT_DASH);
        Assert.assertEquals(ConditionalStyleType.FIRST_ROW, tableStyle.getConditionalStyles().get(0).getType());

        // 3 -  As a property:
        tableStyle.getConditionalStyles().getFirstRow().getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        // Apply padding and text formatting to conditional styles.
        tableStyle.getConditionalStyles().getLastRow().setBottomPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setLeftPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setRightPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setTopPadding(10.0);
        tableStyle.getConditionalStyles().getLastColumn().getFont().setBold(true);

        // List all possible style conditions.
        Iterator<ConditionalStyle> enumerator = tableStyle.getConditionalStyles().iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                ConditionalStyle currentStyle = enumerator.next();
                if (currentStyle != null) System.out.println(currentStyle.getType());
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // Apply the custom style, which contains all conditional styles, to the table.
        table.setStyle(tableStyle);

        // Our style applies some conditional styles by default.
        Assert.assertEquals(TableStyleOptions.FIRST_ROW | TableStyleOptions.FIRST_COLUMN | TableStyleOptions.ROW_BANDS, 
            table.getStyleOptions());

        // We will need to enable all other styles ourselves via the "StyleOptions" property.
        table.setStyleOptions(table.getStyleOptions() | TableStyleOptions.LAST_ROW | TableStyleOptions.LAST_COLUMN);

        doc.save(getArtifactsDir() + "Table.ConditionalStyles.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.ConditionalStyles.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(TableStyleOptions.DEFAULT | TableStyleOptions.LAST_ROW | TableStyleOptions.LAST_COLUMN, table.getStyleOptions());
        ConditionalStyleCollection conditionalStyles = ((TableStyle)doc.getStyles().get("MyTableStyle1")).getConditionalStyles();

        Assert.assertEquals(ConditionalStyleType.FIRST_ROW, conditionalStyles.get(0).getType());
        Assert.assertEquals(msColor.getAliceBlue().getRGB(), conditionalStyles.get(0).getShading().getBackgroundPatternColor().getRGB());
        Assert.assertEquals(Color.BLACK.getRGB(), conditionalStyles.get(0).getBorders().getColor().getRGB());
        Assert.assertEquals(LineStyle.DOT_DASH, conditionalStyles.get(0).getBorders().getLineStyle());
        Assert.assertEquals(ParagraphAlignment.CENTER, conditionalStyles.get(0).getParagraphFormat().getAlignment());

        Assert.assertEquals(ConditionalStyleType.LAST_ROW, conditionalStyles.get(2).getType());
        Assert.assertEquals(10.0d, conditionalStyles.get(2).getBottomPadding());
        Assert.assertEquals(10.0d, conditionalStyles.get(2).getLeftPadding());
        Assert.assertEquals(10.0d, conditionalStyles.get(2).getRightPadding());
        Assert.assertEquals(10.0d, conditionalStyles.get(2).getTopPadding());

        Assert.assertEquals(ConditionalStyleType.LAST_COLUMN, conditionalStyles.get(3).getType());
        Assert.assertTrue(conditionalStyles.get(3).getFont().getBold());
    }

    @Test
    public void clearTableStyleFormatting() throws Exception
    {
        //ExStart
        //ExFor:ConditionalStyle.ClearFormatting
        //ExFor:ConditionalStyleCollection.ClearFormatting
        //ExSummary:Shows how to reset conditional table styles.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("First row");
        builder.endRow();
        builder.insertCell();
        builder.write("Last row");
        builder.endTable();

        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        table.setStyle(tableStyle);

        // Set the table style to color the borders of the first row of the table in red.
        tableStyle.getConditionalStyles().getFirstRow().getBorders().setColor(Color.RED);

        // Set the table style to color the borders of the last row of the table in blue.
        tableStyle.getConditionalStyles().getLastRow().getBorders().setColor(Color.BLUE);

        // Below are two ways of using the "ClearFormatting" method to clear the conditional styles.
        // 1 -  Clear the conditional styles for a specific part of a table:
        tableStyle.getConditionalStyles().get(0).clearFormatting();

        Assert.assertEquals(msColor.Empty, tableStyle.getConditionalStyles().getFirstRow().getBorders().getColor());

        // 2 -  Clear the conditional styles for the entire table:
        tableStyle.getConditionalStyles().clearFormatting();

        Assert.True(tableStyle.getConditionalStyles().All(s => s.Borders.Color == Color.Empty));
        //ExEnd
    }

    @Test
    public void alternatingRowStyles() throws Exception
    {
        //ExStart
        //ExFor:TableStyle.ColumnStripe
        //ExFor:TableStyle.RowStripe
        //ExSummary:Shows how to create conditional table styles that alternate between rows.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // We can configure a conditional style of a table to apply a different color to the row/column,
        // based on whether the row/column is even or odd, creating an alternating color pattern.
        // We can also apply a number n to the row/column banding,
        // meaning that the color alternates after every n rows/columns instead of one.
        // Create a table where single columns and rows will band the columns will banded in threes.
        Table table = builder.startTable();
        for (int i = 0; i < 15; i++)
        {
            for (int j = 0; j < 4; j++)
            {
                builder.insertCell();
                builder.writeln($"{(j % 2 == 0 ? "Even" : "Odd")} column.");
                builder.write($"Row banding {(i % 3 == 0 ? "start" : "continuation")}.");
            }
            builder.endRow();
        }
        builder.endTable();

        // Apply a line style to all the borders of the table.
        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.getBorders().setColor(Color.BLACK);
        tableStyle.getBorders().setLineStyle(LineStyle.DOUBLE);

        // Set the two colors, which will alternate over every 3 rows.
        tableStyle.setRowStripe(3);
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.ODD_ROW_BANDING).getShading().setBackgroundPatternColor(Color.LightBlue);
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.EVEN_ROW_BANDING).getShading().setBackgroundPatternColor(Color.LightCyan);

        // Set a color to apply to every even column, which will override any custom row coloring.
        tableStyle.setColumnStripe(1);
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.EVEN_COLUMN_BANDING).getShading().setBackgroundPatternColor(Color.LightSalmon);

        table.setStyle(tableStyle);

        // The "StyleOptions" property enables row banding by default.
        Assert.assertEquals(TableStyleOptions.FIRST_ROW | TableStyleOptions.FIRST_COLUMN | TableStyleOptions.ROW_BANDS,
            table.getStyleOptions());

        // Use the "StyleOptions" property also to enable column banding.
        table.setStyleOptions(table.getStyleOptions() | TableStyleOptions.COLUMN_BANDS);

        doc.save(getArtifactsDir() + "Table.AlternatingRowStyles.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.AlternatingRowStyles.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);
        tableStyle = (TableStyle)doc.getStyles().get("MyTableStyle1");

        Assert.assertEquals(tableStyle, table.getStyle());
        Assert.assertEquals(table.getStyleOptions() | TableStyleOptions.COLUMN_BANDS, table.getStyleOptions());

        Assert.assertEquals(Color.BLACK.getRGB(), tableStyle.getBorders().getColor().getRGB());
        Assert.assertEquals(LineStyle.DOUBLE, tableStyle.getBorders().getLineStyle());
        Assert.assertEquals(3, tableStyle.getRowStripe());
        Assert.assertEquals(Color.LightBlue.getRGB(), tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.ODD_ROW_BANDING).getShading().getBackgroundPatternColor().getRGB());
        Assert.assertEquals(Color.LightCyan.getRGB(), tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.EVEN_ROW_BANDING).getShading().getBackgroundPatternColor().getRGB());
        Assert.assertEquals(1, tableStyle.getColumnStripe());
        Assert.assertEquals(Color.LightSalmon.getRGB(), tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.EVEN_COLUMN_BANDING).getShading().getBackgroundPatternColor().getRGB());
    }

    @Test
    public void convertToHorizontallyMergedCells() throws Exception
    {
        //ExStart
        //ExFor:Table.ConvertToHorizontallyMergedCells
        //ExSummary:Shows how to convert cells horizontally merged by width to cells merged by CellFormat.HorizontalMerge.
        Document doc = new Document(getMyDir() + "Table with merged cells.docx");

        // Microsoft Word does not write merge flags anymore, defining merged cells by width instead.
        // Aspose.Words by default define only 5 cells in a row, and none of them have the horizontal merge flag,
        // even though there were 7 cells in the row before the horizontal merging took place.
        Table table = doc.getFirstSection().getBody().getTables().get(0);
        Row row = table.getRows().get(0);

        Assert.assertEquals(5, row.getCells().getCount());
        Assert.True(row.getCells().All(c => ((Cell)c).CellFormat.HorizontalMerge == CellMerge.None));

        // Use the "ConvertToHorizontallyMergedCells" method to convert cells horizontally merged
        // by its width to the cell horizontally merged by flags.
        // Now, we have 7 cells, and some of them have horizontal merge values.
        table.convertToHorizontallyMergedCells();
        row = table.getRows().get(0);

        Assert.assertEquals(7, row.getCells().getCount());

        Assert.assertEquals(CellMerge.NONE, row.getCells().get(0).getCellFormat().getHorizontalMerge());
        Assert.assertEquals(CellMerge.FIRST, row.getCells().get(1).getCellFormat().getHorizontalMerge());
        Assert.assertEquals(CellMerge.PREVIOUS, row.getCells().get(2).getCellFormat().getHorizontalMerge());
        Assert.assertEquals(CellMerge.NONE, row.getCells().get(3).getCellFormat().getHorizontalMerge());
        Assert.assertEquals(CellMerge.FIRST, row.getCells().get(4).getCellFormat().getHorizontalMerge());
        Assert.assertEquals(CellMerge.PREVIOUS, row.getCells().get(5).getCellFormat().getHorizontalMerge());
        Assert.assertEquals(CellMerge.NONE, row.getCells().get(6).getCellFormat().getHorizontalMerge());
        //ExEnd
    }
}
