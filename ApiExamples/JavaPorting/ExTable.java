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
import com.aspose.words.TableCollection;
import org.testng.Assert;
import com.aspose.ms.System.msConsole;
import com.aspose.words.RowCollection;
import com.aspose.words.CellCollection;
import com.aspose.ms.System.msString;
import com.aspose.words.SaveFormat;
import com.aspose.words.Table;
import com.aspose.words.NodeType;
import com.aspose.words.Node;
import com.aspose.words.Row;
import com.aspose.words.Cell;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.StoryType;
import com.aspose.words.Section;
import com.aspose.words.AutoFitBehavior;
import com.aspose.words.HeightRule;
import com.aspose.words.TableAlignment;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.BorderType;
import com.aspose.words.LineStyle;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.TextureIndex;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.TextOrientation;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeCollection;
import com.aspose.words.PreferredWidthType;
import com.aspose.words.Run;
import com.aspose.words.CellMerge;
import com.aspose.ms.System.Drawing.msPoint;
import com.aspose.ms.System.Drawing.Rectangle;
import com.aspose.words.TextWrapping;
import com.aspose.words.VerticalAlignment;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.TableStyle;
import com.aspose.words.StyleType;
import com.aspose.words.ConditionalStyleType;
import com.aspose.words.ParagraphAlignment;
import java.util.Iterator;
import com.aspose.words.ConditionalStyle;


/// <summary>
/// Examples using tables in documents.
/// </summary>
@Test
public class ExTable extends ApiExampleBase
{
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
        //ExSummary:Shows how to iterate through all tables in the document and display the content from each cell.
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Here we get all tables from the Document node. You can do this for any other composite node
        // which can contain block level nodes. For example you can retrieve tables from header or from a cell
        // containing another table (nested tables)
        TableCollection tables = doc.getFirstSection().getBody().getTables();

        // We can make a new array to clone all of the tables in the collection
        Assert.assertEquals(2, tables.toArray().length);

        // Iterate through all tables in the document
        for (int i = 0; i < tables.getCount(); i++)
        {
            // Get the index of the table node as contained in the parent node of the table
            System.out.println("Start of Table {i}");

            RowCollection rows = tables.get(i).getRows();

            // RowCollections can be cloned into arrays
            Assert.assertEquals(rows, rows.toArray());
            Assert.assertNotSame(rows, rows.toArray());

            // Iterate through all rows in the table
            for (int j = 0; j < rows.getCount(); j++)
            {
                System.out.println("\tStart of Row {j}");

                CellCollection cells = rows.get(j).getCells();

                // RowCollections can also be cloned into arrays 
                Assert.assertEquals(cells, cells.toArray());
                Assert.assertNotSame(cells, cells.toArray());

                // Iterate through all cells in the row
                for (int k = 0; k < cells.getCount(); k++)
                {
                    // Get the plain text content of this cell
                    String cellText = msString.trim(cells.get(k).toString(SaveFormat.TEXT));
                    // Print the content of the cell
                    System.out.println("\t\tContents of Cell:{k} = \"{cellText}\"");
                }

                System.out.println("\tEnd of Row {j}");
            }

            System.out.println("End of Table {i}\n");
        }
        //ExEnd

        Assert.That(tables.getCount(), Is.GreaterThan(0));
    }

    //ExStart
    //ExFor:Node.GetAncestor(NodeType)
    //ExFor:Node.GetAncestor(System.Type)
    //ExFor:Table.NodeType
    //ExFor:Cell.Tables
    //ExFor:TableCollection
    //ExFor:NodeCollection.Count
    //ExSummary:Shows how to find out if a table contains another table or if the table itself is nested inside another table.
    @Test //ExSkip
    public void calculateDepthOfNestedTables() throws Exception
    {
        Document doc = new Document(getMyDir() + "Nested tables.docx");
        int tableIndex = 0;

        for (Table table : doc.getChildNodes(NodeType.TABLE, true).<Table>OfType() !!Autoporter error: Undefined expression type )
        {
            // First lets find if any cells in the table have tables themselves as children
            int count = getChildTableCount(table);
            msConsole.writeLine("Table #{0} has {1} tables directly within its cells", tableIndex, count);

            // Now let's try the other way around, lets try find if the table is nested inside another table and at what depth
            int tableDepth = getNestedDepthOfTable(table);

            if (tableDepth > 0)
                msConsole.writeLine("Table #{0} is nested inside another table at depth of {1}", tableIndex,
                    tableDepth);
            else
                msConsole.writeLine("Table #{0} is a non nested table (is not a child of another table)", tableIndex);

            tableIndex++;
        }
    }

    /// <summary>
    /// Calculates what level a table is nested inside other tables.
    /// <returns>
    /// An integer containing the level the table is nested at.
    /// 0 = Table is not nested inside any other table
    /// 1 = Table is nested within one parent table
    /// 2 = Table is nested within two parent tables etc..</returns>
    /// </summary>
    private static int getNestedDepthOfTable(Table table)
    {
        int depth = 0;

        // The parent of the table will be a Cell, instead attempt to find a grandparent that is of type Table
        Node parent = table.getAncestor(table.getNodeType());

        while (parent != null)
        {
            // Every time we find a table a level up we increase the depth counter and then try to find an
            // ancestor of type table from the parent
            depth++;
            parent = parent.getAncestor(Table.class);
        }

        return depth;
    }

    /// <summary>
    /// Determines if a table contains any immediate child table within its cells.
    /// Does not recursively traverse through those tables to check for further tables.
    /// <returns>Returns true if at least one child cell contains a table.
    /// Returns false if no cells in the table contains a table.</returns>
    /// </summary>
    private static int getChildTableCount(Table table)
    {
        int tableCount = 0;
        // Iterate through all child rows in the table
        for (Row row : table.getRows().<Row>OfType() !!Autoporter error: Undefined expression type )
        {
            // Iterate through all child cells in the row
            for (Cell Cell : row.getCells().<Cell>OfType() !!Autoporter error: Undefined expression type )
            {
                // Retrieve the collection of child tables of this cell
                TableCollection childTables = Cell.getTables();

                // If this cell has a table as a child then return true
                if (childTables.getCount() > 0)
                    tableCount++;
            }
        }

        // No cell contains a table
        return tableCount;
    }
    //ExEnd

    @Test
    public void convertTextBoxToTable() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box
        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 300.0, 50.0);

        // Move the builder into the text box and write text
        builder.moveTo(textBox.getLastParagraph());
        builder.write("Hello world!");

        // Convert all shape nodes which contain child nodes
        // We convert the collection to an array as static "snapshot" because the original textboxes will be removed after conversion which will
        // invalidate the enumerator
        for (Shape shape : doc.getChildNodes(NodeType.SHAPE, true).toArray().<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            if (shape.hasChildNodes())
            {
                convertTextboxToTable(shape);
            }
        }

        doc.save(getArtifactsDir() + "Table.ConvertTextBoxToTable.html");
    }

    /// <summary>
    /// Converts a textbox to a table by copying the same content and formatting.
    /// Currently export to HTML will render the textbox as an image which looses any text functionality.
    /// This is useful to convert textboxes in order to retain proper text.
    /// </summary>
    /// <param name="textBox">The textbox shape to convert to a table</param>
    private static void convertTextboxToTable(Shape textBox) throws Exception
    {
        if (textBox.getStoryType() != StoryType.TEXTBOX)
            throw new IllegalArgumentException("Can only convert a shape of type textbox");

        Document doc = (Document) textBox.getDocument();
        Section section = (Section) textBox.getAncestor(NodeType.SECTION);

        // Create a table to replace the textbox and transfer the same content and formatting
        Table table = new Table(doc);
        // Ensure that the table contains a row and a cell
        table.ensureMinimum();
        // Use fixed column widths
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        // A shape is inline level (within a paragraph) where a table can only be block level so insert the table
        // after the paragraph which contains the shape
        Node shapeParent = textBox.getParentNode();
        shapeParent.getParentNode().insertAfter(table, shapeParent);

        // If the textbox is not inline then try to match the shape's left position using the table's left indent
        if (!textBox.isInline() && textBox.getLeft() < section.getPageSetup().getPageWidth())
            table.setLeftIndent(textBox.getLeft());

        // We are only using one cell to replicate a textbox so we can make use of the FirstRow and FirstCell property
        // Carry over borders and shading
        Row firstRow = table.getFirstRow();
        Cell firstCell = firstRow.getFirstCell();
        firstCell.getCellFormat().getBorders().setColor(textBox.getStrokeColor());
        firstCell.getCellFormat().getBorders().setLineWidth(textBox.getStrokeWeight());
        firstCell.getCellFormat().getShading().setBackgroundPatternColor(textBox.getFill().getColor());

        // Transfer the same height and width of the textbox to the table
        firstRow.getRowFormat().setHeightRule(HeightRule.EXACTLY);
        firstRow.getRowFormat().setHeight(textBox.getHeight());
        firstCell.getCellFormat().setWidth(textBox.getWidth());
        table.setAllowAutoFit(false);

        // Replicate the textbox's horizontal alignment
        /*TableAlignment*/int horizontalAlignment;
        switch (textBox.getHorizontalAlignment())
        {
            case HorizontalAlignment.LEFT:
                horizontalAlignment = TableAlignment.LEFT;
                break;
            case HorizontalAlignment.CENTER:
                horizontalAlignment = TableAlignment.CENTER;
                break;
            case HorizontalAlignment.RIGHT:
                horizontalAlignment = TableAlignment.RIGHT;
                break;
            default:
                // Most other options are left by default
                horizontalAlignment = TableAlignment.LEFT;
                break;
        }

        table.setAlignment(horizontalAlignment);
        firstCell.removeAllChildren();

        // Append all content from the textbox to the new table
        for (Node node : textBox.getChildNodes(NodeType.ANY, false).toArray())
        {
            table.getFirstRow().getFirstCell().appendChild(node);
        }

        // Remove the empty textbox from the document
        textBox.remove();
    }

    @Test
    public void ensureTableMinimum() throws Exception
    {
        //ExStart
        //ExFor:Table.EnsureMinimum
        //ExSummary:Shows how to ensure a table node is valid.
        Document doc = new Document();

        // Create a new table and add it to the document
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);

        // Ensure the table is valid (has at least one row with one cell)
        table.ensureMinimum();
        //ExEnd
    }

    @Test
    public void ensureRowMinimum() throws Exception
    {
        //ExStart
        //ExFor:Row.EnsureMinimum
        //ExSummary:Shows how to ensure a row node is valid.
        Document doc = new Document();

        // Create a new table and add it to the document
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);

        // Create a new row and add it to the table
        Row row = new Row(doc);
        table.appendChild(row);

        // Ensure the row is valid (has at least one cell)
        row.ensureMinimum();
        //ExEnd
    }

    @Test
    public void ensureCellMinimum() throws Exception
    {
        //ExStart
        //ExFor:Cell.EnsureMinimum
        //ExSummary:Shows how to ensure a cell node is valid.
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Gets the first cell in the document
        Cell cell = (Cell) doc.getChild(NodeType.CELL, 0, true);

        // Ensure the cell is valid (the last child is a paragraph)
        cell.ensureMinimum();
        //ExEnd
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
        //ExSummary:Shows how to apply a outline border to a table.
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Align the table to the center of the page
        table.setAlignment(TableAlignment.CENTER);

        // Clear any existing borders and shading from the table
        table.clearBorders();
        table.clearShading();

        // Set a green border around the table but not inside
        table.setBorder(BorderType.LEFT, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.RIGHT, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.TOP, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.BOTTOM, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);

        // Fill the cells with a light green solid color
        table.setShading(TextureIndex.TEXTURE_SOLID, msColor.getLightGreen(), msColor.Empty);

        doc.save(getArtifactsDir() + "Table.SetOutlineBorders.docx");
        //ExEnd

        // Verify the borders were set correctly
        Assert.assertEquals(TableAlignment.CENTER, table.getAlignment());
        Assert.assertEquals(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getTop().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getLeft().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getRight().getColor().getRGB());
        Assert.assertEquals(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getBottom().getColor().getRGB());
        msAssert.areNotEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getHorizontal().getColor().getRGB());
        msAssert.areNotEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getVertical().getColor().getRGB());
        Assert.assertEquals(msColor.getLightGreen().getRGB(),
            table.getFirstRow().getFirstCell().getCellFormat().getShading().getForegroundPatternColor().getRGB());
    }

    @Test
    public void setTableBorders() throws Exception
    {
        //ExStart
        //ExFor:Table.SetBorders
        //ExSummary:Shows how to build a table with all borders enabled (grid).
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Clear any existing borders from the table
        table.clearBorders();

        // Set a green border around and inside the table
        table.setBorders(LineStyle.SINGLE, 1.5, msColor.getGreen());

        doc.save(getArtifactsDir() + "Table.SetAllBorders.doc");
        //ExEnd

        // Verify the borders were set correctly
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
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Retrieve the first row in the table
        Row firstRow = table.getFirstRow();

        // Modify some row level properties
        firstRow.getRowFormat().getBorders().setLineStyle(LineStyle.NONE);
        firstRow.getRowFormat().setHeightRule(HeightRule.AUTO);
        firstRow.getRowFormat().setAllowBreakAcrossPages(true);
        //ExEnd

        doc.save(getArtifactsDir() + "Table.RowFormat.doc");

        doc = new Document(getArtifactsDir() + "Table.RowFormat.doc");
        table = (Table)doc.getChild(NodeType.TABLE, 0, true);
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
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Retrieve the first cell in the table
        Cell firstCell = table.getFirstRow().getFirstCell();

        // Modify some row level properties
        firstCell.getCellFormat().setWidth(30.0); // in points
        firstCell.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        firstCell.getCellFormat().getShading().setForegroundPatternColor(msColor.getLightGreen());
        //ExEnd

        doc.save(getArtifactsDir() + "Table.CellFormat.doc");

        doc = new Document(getArtifactsDir() + "Table.CellFormat.doc");
        table = (Table)doc.getChild(NodeType.TABLE, 0, true);
        Assert.assertEquals(30, table.getFirstRow().getFirstCell().getCellFormat().getWidth());
        Assert.assertEquals(TextOrientation.DOWNWARD, table.getFirstRow().getFirstCell().getCellFormat().getOrientation());
        Assert.assertEquals(msColor.getLightGreen().getRGB(),
            table.getFirstRow().getFirstCell().getCellFormat().getShading().getForegroundPatternColor().getRGB());
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

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(25.9d, table.getDistanceTop());
        Assert.assertEquals(25.9d, table.getDistanceBottom());
        Assert.assertEquals(17.3d, table.getDistanceLeft());
        Assert.assertEquals(17.3d, table.getDistanceRight());
        //ExEnd
    }

    @Test
    public void clearBorders() throws Exception
    {
        //ExStart
        //ExFor:Table.ClearBorders
        //ExSummary:Shows how to remove all borders from a table.
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Remove all borders from the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Clear the borders all cells in the table
        table.clearBorders();

        doc.save(getArtifactsDir() + "Table.ClearBorders.doc");
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

        // Create a table and give it conditional styling on border colors based on the row being the first or last
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Carrots");
        builder.insertCell();
        builder.write("30");
        builder.endRow();
        builder.insertCell();
        builder.write("Potatoes");
        builder.insertCell();
        builder.write("50");
        builder.endTable();

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(true);
        options.setFindWholeWordsOnly(true);

        // Replace any instances of our String in the entire table
        table.getRange().replace("Carrots", "Eggs", options);
        // Replace any instances of our String in the last cell of the table only
        table.getLastRow().getLastCell().getRange().replace("50", "20", options);

        doc.save(getArtifactsDir() + "Table.ReplaceCellText.doc");
        //ExEnd

        Assert.assertEquals("20", msString.trim(table.getLastRow().getLastCell().toString(SaveFormat.TEXT)));
    }

    @Test
    public void printTableRange() throws Exception
    {
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Get the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // The range text will include control characters such as "\a" for a cell
        // You can call ToString on the desired node to retrieve the plain text content

        // Print the plain text range of the table to the screen
        System.out.println("Contents of the table: ");
        System.out.println(table.getRange().getText());
        
        // Print the contents of the second row to the screen
        System.out.println("\nContents of the row: ");
        System.out.println(table.getRows().get(1).getRange().getText());

        // Print the contents of the last cell in the table to the screen
        System.out.println("\nContents of the cell: ");
        System.out.println(table.getLastRow().getLastCell().getRange().getText());
        
        Assert.assertEquals("\u0007Column 1\u0007Column 2\u0007Column 3\u0007Column 4\u0007\u0007", table.getRows().get(1).getRange().getText());
        Assert.assertEquals("Cell 12 contents\u0007", table.getLastRow().getLastCell().getRange().getText());
    }

    @Test
    public void cloneTable() throws Exception
    {
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Retrieve the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Create a clone of the table
        Table tableClone = (Table) table.deepClone(true);

        // Insert the cloned table into the document after the original
        table.getParentNode().insertAfter(tableClone, table);

        // Insert an empty paragraph between the two tables or else they will be combined into one
        // upon save. This has to do with document validation
        table.getParentNode().insertAfter(new Paragraph(doc), table);

        doc.save(getArtifactsDir() + "Table.CloneTable.doc");
        
        // Verify that the table was cloned and inserted properly
        Assert.assertEquals(3, doc.getChildNodes(NodeType.TABLE, true).getCount());
        Assert.assertEquals(table.getRange().getText(), tableClone.getRange().getText());

        for (Cell cell : tableClone.getChildNodes(NodeType.CELL, true).<Cell>OfType() !!Autoporter error: Undefined expression type )
            cell.removeAllChildren();
        
        Assert.assertEquals("", msString.trim(tableClone.toString(SaveFormat.TEXT)));
    }

    @Test
    public void disableBreakAcrossPages() throws Exception
    {
        //ExStart
        //ExFor:RowFormat.AllowBreakAcrossPages
        //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
        // Disable breaking across pages for all rows in the table
        Document doc = new Document(getMyDir() + "Table spanning two pages.docx");

        // Retrieve the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        for (Row row : table.<Row>OfType() !!Autoporter error: Undefined expression type )
            row.getRowFormat().setAllowBreakAcrossPages(false);

        doc.save(getArtifactsDir() + "Table.DisableBreakAcrossPages.docx");
        //ExEnd

        Assert.assertFalse(table.getFirstRow().getRowFormat().getAllowBreakAcrossPages());
        Assert.assertFalse(table.getLastRow().getRowFormat().getAllowBreakAcrossPages());
    }

    @Test
    public void allowAutoFitOnTable() throws Exception
    {
        Document doc = new Document();

        Table table = new Table(doc);
        table.ensureMinimum();

        //ExStart
        //ExFor:Table.AllowAutoFit
        //ExSummary:Shows how to set a table to shrink or grow each cell to accommodate its contents.
        table.setAllowAutoFit(true);
        //ExEnd
    }

    @Test
    public void keepTableTogether() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table spanning two pages.docx");

        // Retrieve the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        //ExStart
        //ExFor:ParagraphFormat.KeepWithNext
        //ExFor:Row.IsLastRow
        //ExFor:Paragraph.IsEndOfCell
        //ExFor:Paragraph.IsInCell
        //ExFor:Cell.ParentRow
        //ExFor:Cell.Paragraphs
        //ExSummary:Shows how to set a table to stay together on the same page.
        // To keep a table from breaking across a page we need to enable KeepWithNext 
        // for every paragraph in the table except for the last paragraphs in the last 
        // row of the table
        for (Cell cell : table.getChildNodes(NodeType.CELL, true).<Cell>OfType() !!Autoporter error: Undefined expression type )
        for (Paragraph para : cell.getParagraphs().<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            // Every paragraph that's inside a cell will have this flag set
            Assert.assertTrue(para.isInCell());

            if (!(cell.getParentRow().isLastRow() && para.isEndOfCell()))
                para.getParagraphFormat().setKeepWithNext(true);
        }
        //ExEnd

        doc.save(getArtifactsDir() + "Table.KeepTableTogether.doc");

        // Verify the correct paragraphs were set properly
        for (Paragraph para : table.getChildNodes(NodeType.PARAGRAPH, true).<Paragraph>OfType() !!Autoporter error: Undefined expression type )
            if (para.isEndOfCell() && ((Cell) para.getParentNode()).getParentRow().isLastRow())
                Assert.assertFalse(para.getParagraphFormat().getKeepWithNext());
            else
                Assert.assertTrue(para.getParagraphFormat().getKeepWithNext());
    }

    @Test
    public void addClonedRowToTable() throws Exception
    {
        //ExStart
        //ExFor:Row
        //ExSummary:Shows how to make a clone of the last row of a table and append it to the table.
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Retrieve the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Clone the last row in the table
        Row clonedRow = (Row) table.getLastRow().deepClone(true);

        // Remove all content from the cloned row's cells. This makes the row ready for
        // new content to be inserted into
        for (Cell cell : clonedRow.getCells().<Cell>OfType() !!Autoporter error: Undefined expression type )
            cell.removeAllChildren();

        // Add the row to the end of the table
        table.appendChild(clonedRow);

        doc.save(getArtifactsDir() + "Table.AddCloneRowToTable.doc");
        //ExEnd

        // Verify that the row was cloned and appended properly
        Assert.assertEquals(6, table.getRows().getCount());
        Assert.assertEquals("", msString.trim(table.getLastRow().toString(SaveFormat.TEXT)));
        Assert.assertEquals(5, table.getLastRow().getCells().getCount());
    }

    @Test
    public void fixDefaultTableWidthsInAw105() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Keep a reference to the table being built
        Table table = builder.startTable();

        // Apply some formatting
        builder.getCellFormat().setWidth(100.0);
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.RED);

        builder.insertCell();
        // This will cause the table to be structured using column widths as in previous versions
        // instead of fitted to the page width like in the newer versions
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        // Continue with building your table as usual...
    }

    @Test
    public void fixDefaultTableBordersIn105() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Keep a reference to the table being built
        Table table = builder.startTable();

        builder.insertCell();
        // Clear all borders to match the defaults used in previous versions
        table.clearBorders();

        // Continue with building your table as usual...
    }

    @Test
    public void fixDefaultTableFormattingExceptionIn105() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Keep a reference to the table being built
        Table table = builder.startTable();

        // We must first insert a new cell which in turn inserts a row into the table
        builder.insertCell();
        // Once a row exists in our table we can apply table wide formatting
        table.setAllowAutoFit(true);

        // Continue with building your table as usual...
    }

    @Test
    public void fixRowFormattingNotAppliedIn105() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();

        // For the first row this will be set correctly
        builder.getRowFormat().setHeadingFormat(true);

        builder.insertCell();
        builder.writeln("Text");
        builder.insertCell();
        builder.writeln("Text");

        // End the first row
        builder.endRow();

        // Here we would normally define some other row formatting, such as disabling the 
        // heading format. However at the moment this will be ignored and the value from the 
        // first row reapplied to the row

        builder.insertCell();

        // Instead make sure to specify the row formatting for the second row here
        builder.getRowFormat().setHeadingFormat(false);

        // Continue with building your table as usual...
    }

    @Test
    public void getIndexOfTableElements() throws Exception
    {
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        //ExStart
        //ExFor:NodeCollection.IndexOf(Node)
        //ExSummary:Retrieves the index of a table in the document.
        NodeCollection allTables = doc.getChildNodes(NodeType.TABLE, true);
        int tableIndex = allTables.indexOf(table);

        Row row = table.getRows().get(2);
        int rowIndex = table.indexOf(row);

        Cell cell = row.getLastCell();
        int cellIndex = row.indexOf(cell);
        //ExEnd

        Assert.assertEquals(0, tableIndex);
        Assert.assertEquals(2, rowIndex);
        Assert.assertEquals(4, cellIndex);
    }

    @Test
    public void getPreferredWidthTypeAndValue() throws Exception
    {
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Find the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        //ExStart
        //ExFor:PreferredWidthType
        //ExFor:PreferredWidth.Type
        //ExFor:PreferredWidth.Value
        //ExSummary:Retrieves the preferred width type of a table cell.
        Cell firstCell = table.getFirstRow().getFirstCell();
        /*PreferredWidthType*/int type = firstCell.getCellFormat().getPreferredWidth().getType();
        double value = firstCell.getCellFormat().getPreferredWidth().getValue();
        //ExEnd

        Assert.assertEquals(PreferredWidthType.PERCENT, type);
        Assert.assertEquals(11.16, value);
    }

    @Test
    public void insertTableUsingNodes() throws Exception
    {
        //ExStart
        //ExFor:Table.AllowCellSpacing
        //ExFor:Row
        //ExFor:Row.RowFormat
        //ExFor:RowFormat
        //ExFor:Cell.CellFormat
        //ExFor:CellFormat
        //ExFor:CellFormat.Shading
        //ExFor:Cell.FirstParagraph
        //ExSummary:Shows how to insert a table using the constructors of nodes.
        Document doc = new Document();

        // We start by creating the table object. Note how we must pass the document object
        // to the constructor of each node. This is because every node we create must belong
        // to some document
        Table table = new Table(doc);
        // Add the table to the document
        doc.getFirstSection().getBody().appendChild(table);

        // Here we could call EnsureMinimum to create the rows and cells for us. This method is used
        // to ensure that the specified node is valid, in this case a valid table should have at least one
        // row and one cell, therefore this method creates them for us

        // Instead we will handle creating the row and table ourselves. This would be the best way to do this
        // if we were creating a table inside an algorithm for example
        Row row = new Row(doc);
        row.getRowFormat().setAllowBreakAcrossPages(true);
        table.appendChild(row);

        // We can now apply any auto fit settings
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        // Create a cell and add it to the row
        Cell cell = new Cell(doc);
        cell.getCellFormat().getShading().setBackgroundPatternColor(Color.LightBlue);
        cell.getCellFormat().setWidth(80.0);

        // Add a paragraph to the cell as well as a new run with some text
        cell.appendChild(new Paragraph(doc));
        cell.getFirstParagraph().appendChild(new Run(doc, "Row 1, Cell 1 Text"));

        // Add the cell to the row
        row.appendChild(cell);

        // We would then repeat the process for the other cells and rows in the table
        // We can also speed things up by cloning existing cells and rows
        row.appendChild(cell.deepClone(false));
        row.getLastCell().appendChild(new Paragraph(doc));
        row.getLastCell().getFirstParagraph().appendChild(new Run(doc, "Row 1, Cell 2 Text"));

        // Remove spacing between cells
        table.setAllowCellSpacing(false);

        doc.save(getArtifactsDir() + "Table.InsertTableUsingNodes.doc");
        //ExEnd

        Assert.assertEquals(1, doc.getChildNodes(NodeType.TABLE, true).getCount());
        Assert.assertEquals(1, doc.getChildNodes(NodeType.ROW, true).getCount());
        Assert.assertEquals(2, doc.getChildNodes(NodeType.CELL, true).getCount());
        Assert.assertEquals("Row 1, Cell 1 Text\r\nRow 1, Cell 2 Text",
            msString.trim(doc.getFirstSection().getBody().getTables().get(0).toString(SaveFormat.TEXT)));
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
    //ExSummary:Shows how to build a nested table without using DocumentBuilder.
    @Test //ExSkip
    public void createNestedTable() throws Exception
    {
        Document doc = new Document();

        // Create the outer table with three rows and four columns
        Table outerTable = createTable(doc, 3, 4, "Outer Table");
        // Add it to the document body
        doc.getFirstSection().getBody().appendChild(outerTable);

        // Create another table with two rows and two columns
        Table innerTable = createTable(doc, 2, 2, "Inner Table");
        // Add this table to the first cell of the outer table
        outerTable.getFirstRow().getFirstCell().appendChild(innerTable);

        doc.save(getArtifactsDir() + "Table.CreateNestedTable.doc");

        Assert.assertEquals(2, doc.getChildNodes(NodeType.TABLE, true).getCount()); // ExSkip
        Assert.assertEquals(1, outerTable.getFirstRow().getFirstCell().getTables().getCount()); //ExSkip
        Assert.assertEquals(16, outerTable.getChildNodes(NodeType.CELL, true).getCount()); //ExSkip
        Assert.assertEquals(4, innerTable.getChildNodes(NodeType.CELL, true).getCount()); //ExSkip
        Assert.assertEquals("Aspose table title", innerTable.getTitle()); //ExSkip
        Assert.assertEquals("Aspose table description", innerTable.getDescription()); //ExSkip
    }

    /// <summary>
    /// Creates a new table in the document with the given dimensions and text in each cell.
    /// </summary>
    private static Table createTable(Document doc, int rowCount, int cellCount, String cellText) throws Exception
    {
        Table table = new Table(doc);

        // Create the specified number of rows
        for (int rowId = 1; rowId <= rowCount; rowId++)
        {
            Row row = new Row(doc);
            table.appendChild(row);

            // Create the specified number of cells for each row
            for (int cellId = 1; cellId <= cellCount; cellId++)
            {
                Cell cell = new Cell(doc);
                row.appendChild(cell);
                // Add a blank paragraph to the cell
                cell.appendChild(new Paragraph(doc));

                // Add the text
                cell.getFirstParagraph().appendChild(new Run(doc, cellText));
            }
        }

        // You can add title and description to your table only when added at least one row to the table first
        // This properties are meaningful for ISO / IEC 29500 compliant DOCX documents(see the OoxmlCompliance class)
        // When saved to pre-ISO/IEC 29500 formats, the properties are ignored
        table.setTitle("Aspose table title");
        table.setDescription("Aspose table description");

        return table;
    }
    //ExEnd

    //ExStart
    //ExFor:CellFormat.HorizontalMerge
    //ExFor:CellFormat.VerticalMerge
    //ExFor:CellMerge
    //ExSummary:Prints the horizontal and vertical merge type of a cell.
    @Test //ExSkip
    public void checkCellsMerged() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table with merged cells.docx");

        // Retrieve the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        for (Row row : table.getRows().<Row>OfType() !!Autoporter error: Undefined expression type )
        {
            for (Cell cell : row.getCells().<Cell>OfType() !!Autoporter error: Undefined expression type )
            {
                System.out.println(printCellMergeType(cell));
            }
        }

        Assert.assertEquals("The cell at R1, C1 is vertically merged",
            printCellMergeType(table.getFirstRow().getFirstCell())); //ExSkip
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
        // Open the document
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Retrieve the first table in the body of the first section
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        // We want to merge the range of cells found in between these two cells
        Cell cellStartRange = table.getRows().get(2).getCells().get(2);
        Cell cellEndRange = table.getRows().get(3).getCells().get(3);

        // Merge all the cells between the two specified cells into one
        mergeCells(cellStartRange, cellEndRange);

        // Save the document
        doc.save(getArtifactsDir() + "Table.MergeCellRange.doc");

        // Verify the cells were merged
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
    /// Merges the range of cells found between the two specified cells both horizontally and vertically. Can span over multiple rows.
    /// </summary>
    @Test (enabled = false)
    public static void mergeCells(Cell startCell, Cell endCell)
    {
        Table parentTable = startCell.getParentRow().getParentTable();

        // Find the row and cell indices for the start and end cell
        /*Point*/long startCellPos = msPoint.ctor(startCell.getParentRow().indexOf(startCell),
            parentTable.indexOf(startCell.getParentRow()));
        /*Point*/long endCellPos = msPoint.ctor(endCell.getParentRow().indexOf(endCell), parentTable.indexOf(endCell.getParentRow()));
        // Create the range of cells to be merged based off these indices
        // Inverse each index if the end cell if before the start cell
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
                // Check if the current cell is inside our merge range then merge it
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
        // Load the document
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Get the first and second table in the document
        // The rows from the second table will be appended to the end of the first table
        Table firstTable = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Table secondTable = (Table) doc.getChild(NodeType.TABLE, 1, true);

        // Append all rows from the current table to the next
        // Due to the design of tables even tables with different cell count and widths can be joined into one table
        while (secondTable.hasChildNodes())
            firstTable.getRows().add(secondTable.getFirstRow());

        // Remove the empty table container
        secondTable.remove();

        doc.save(getArtifactsDir() + "Table.CombineTables.doc");
        //ExEnd

        Assert.assertEquals(1, doc.getChildNodes(NodeType.TABLE, true).getCount());
        Assert.assertEquals(9, doc.getFirstSection().getBody().getTables().get(0).getRows().getCount());
        Assert.assertEquals(42, doc.getFirstSection().getBody().getTables().get(0).getChildNodes(NodeType.CELL, true).getCount());
    }

    @Test
    public void splitTable() throws Exception
    {
        // Load the document
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Get the first table in the document
        Table firstTable = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // We will split the table at the third row (inclusive)
        Row row = firstTable.getRows().get(2);

        // Create a new container for the split table
        Table table = (Table) firstTable.deepClone(false);

        // Insert the container after the original
        firstTable.getParentNode().insertAfter(table, firstTable);

        // Add a buffer paragraph to ensure the tables stay apart
        firstTable.getParentNode().insertAfter(new Paragraph(doc), firstTable);

        Row currentRow;

        do
        {
            currentRow = firstTable.getLastRow();
            table.prependChild(currentRow);
        } while (currentRow != row);

        doc.save(getArtifactsDir() + "Table.SplitTable.doc");

        doc = new Document(getArtifactsDir() + "Table.SplitTable.doc");
        // Test we are adding the rows in the correct order and the 
        // selected row was also moved
        Assert.assertEquals(row, table.getFirstRow());

        Assert.assertEquals(2, firstTable.getRows().getCount());
        Assert.assertEquals(3, table.getRows().getCount());
        Assert.assertEquals(3, doc.getChildNodes(NodeType.TABLE, true).getCount());
    }

    @Test
    public void checkDefaultValuesForFloatingTableProperties() throws Exception
    {
        //ExStart
        //ExFor:Table.TextWrapping
        //ExFor:TextWrapping
        //ExSummary:Shows how to work with table text wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = DocumentHelper.insertTable(builder);

        if (table.getTextWrapping() == TextWrapping.AROUND)
        {
            Assert.assertEquals(HorizontalAlignment.DEFAULT, table.getRelativeHorizontalAlignment());
            Assert.assertEquals(VerticalAlignment.DEFAULT, table.getRelativeVerticalAlignment());
            Assert.assertEquals(RelativeHorizontalPosition.COLUMN, table.getHorizontalAnchor());
            Assert.assertEquals(RelativeVerticalPosition.MARGIN, table.getVerticalAnchor());
            Assert.assertEquals(0, table.getAbsoluteHorizontalDistance());
            Assert.assertEquals(0, table.getAbsoluteVerticalDistance());
            Assert.assertEquals(true, table.getAllowOverlap());
        }
        //ExEnd
    }

    @Test
    public void getFloatingTableProperties() throws Exception
    {
        //ExStart
        //ExFor:Table.HorizontalAnchor
        //ExFor:Table.VerticalAnchor
        //ExFor:Table.AllowOverlap
        //ExFor:ShapeBase.AllowOverlap
        //ExSummary:Shows how get properties for floating tables
        Document doc = new Document(getMyDir() + "Table wrapped by text.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        if (table.getTextWrapping() == TextWrapping.AROUND)
        {
            Assert.assertEquals(RelativeHorizontalPosition.MARGIN, table.getHorizontalAnchor());
            Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, table.getVerticalAnchor());
            Assert.assertEquals(false, table.getAllowOverlap());
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
        //ExSummary:Shows how get/set properties for floating tables.
        Document doc = new Document(getMyDir() + "Table wrapped by text.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        table.setAbsoluteHorizontalDistance(10.0);
        table.setAbsoluteVerticalDistance(15.0);

        // Check that absolute distance was set correct
        Assert.assertEquals(10, table.getAbsoluteHorizontalDistance());
        Assert.assertEquals(15, table.getAbsoluteVerticalDistance());

        // Setting RelativeHorizontalAlignment will reset AbsoluteHorizontalDistance to default value and vice versa,
        // the same is for vertical positioning
        table.setRelativeVerticalAlignment(VerticalAlignment.TOP);
        table.setRelativeHorizontalAlignment(HorizontalAlignment.CENTER);
        
        // Check that AbsoluteHorizontalDistance and AbsoluteVerticalDistance are reset 
        Assert.assertEquals(0, table.getAbsoluteHorizontalDistance());
        Assert.assertEquals(0, table.getAbsoluteVerticalDistance());
        Assert.assertEquals(VerticalAlignment.TOP, table.getRelativeVerticalAlignment());
        Assert.assertEquals(HorizontalAlignment.CENTER, table.getRelativeHorizontalAlignment());

        doc.save(getArtifactsDir() + "Table.ChangeFloatingTableProperties.docx");
        //ExEnd
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
        //ExSummary:Shows how to create your own style settings for the table.
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
        tableStyle.getBorders().setColor(Color.BLACK);
        tableStyle.getBorders().setLineStyle(LineStyle.DOT_DASH);

        table.setStyle(tableStyle);

        // Some Table attributes are linked to style variables
        Assert.assertEquals(true, table.getBidi());
        Assert.assertEquals(5.0, table.getCellSpacing());
        Assert.assertEquals("MyTableStyle1", table.getStyleName());

        doc.save(getArtifactsDir() + "Table.TableStyleCreation.docx");
        //ExEnd
    }

    @Test
    public void setTableAligment() throws Exception
    {
        //ExStart
        //ExFor:TableStyle.Alignment
        //ExFor:TableStyle.LeftIndent
        //ExSummary:Shows how to set table position.
        Document doc = new Document();
 
        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        // By default AW uses Alignment instead of LeftIndent
        // To set table position use
        tableStyle.setAlignment(TableAlignment.CENTER);
        // or
        tableStyle.setLeftIndent(55.0);
        //ExEnd
    }

    @Test
    public void workWithTableConditionalStyles() throws Exception
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
        //ExFor:ConditionalStyleCollection.Item(TableStyleOverrideType)
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

        // Create a table, which we will partially style
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Cell 1, to be formatted");
        builder.insertCell();
        builder.write("Cell 2, to be formatted");
        builder.endRow();
        builder.insertCell();
        builder.write("Cell 3, to be left unformatted");
        builder.insertCell();
        builder.write("Cell 4, to be left unformatted");
        builder.endTable();

        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        // There is a different ways how to get conditional styles:
        // by conditional style type
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.FIRST_ROW).getShading().setBackgroundPatternColor(msColor.getAliceBlue());
        // by index
        tableStyle.getConditionalStyles().get(0).getBorders().setColor(Color.BLACK);
        tableStyle.getConditionalStyles().get(0).getBorders().setLineStyle(LineStyle.DOT_DASH);
        Assert.assertEquals(ConditionalStyleType.FIRST_ROW, tableStyle.getConditionalStyles().get(0).getType());
        // directly from ConditionalStyleCollection
        tableStyle.getConditionalStyles().getFirstRow().getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        // To see this in Word document select Total Row checkbox in Design Tab
        tableStyle.getConditionalStyles().getLastRow().setBottomPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setLeftPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setRightPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setTopPadding(10.0);
        // To see this in Word document select Last Column checkbox in Design Tab
        tableStyle.getConditionalStyles().getLastColumn().getFont().setBold(true);

        // List all possible style conditions
        Iterator<ConditionalStyle> enumerator = tableStyle.getConditionalStyles().iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                ConditionalStyle currentStyle = enumerator.next();
                if (currentStyle != null) msConsole.writeLine(currentStyle.getType());
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // Apply conditional style to the table and save
        table.setStyle(tableStyle);
        
        doc.save(getArtifactsDir() + "Table.WorkWithTableConditionalStyles.docx");
        //ExEnd
    }

    @Test
    public void clearTableStyleFormatting() throws Exception
    {
        //ExStart
        //ExFor:ConditionalStyle.ClearFormatting
        //ExFor:ConditionalStyleCollection.ClearFormatting
        //ExSummary:Shows how to reset all table styles.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a table and give it conditional styling on border colors based on the row being the first or last
        builder.startTable();
        builder.insertCell();
        builder.write("First row");
        builder.endRow();
        builder.insertCell();
        builder.write("Last row");
        builder.endTable();

        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.getConditionalStyles().getFirstRow().getBorders().setColor(Color.RED);
        tableStyle.getConditionalStyles().getLastRow().getBorders().setColor(Color.BLUE);

        // You can reset styles from the specific table area
        tableStyle.getConditionalStyles().get(0).clearFormatting();
        Assert.assertEquals(msColor.Empty, tableStyle.getConditionalStyles().getFirstRow().getBorders().getColor());

        // Or clear all table styles
        tableStyle.getConditionalStyles().clearFormatting();
        Assert.assertEquals(msColor.Empty, tableStyle.getConditionalStyles().getLastRow().getBorders().getColor());
        //ExEnd
    }

    @Test
    public void workWithOddEvenRowColumnStyles() throws Exception
    {
        //ExStart
        //ExFor:TableStyle.ColumnStripe
        //ExFor:TableStyle.RowStripe
        //ExSummary:Shows how to work with odd/even row/column styles.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a table and give it conditional styling on border colors based on row number parity
        builder.startTable();
        builder.insertCell();
        builder.write("Odd row");
        builder.endRow();
        builder.insertCell();
        builder.write("Even row");
        builder.endRow();
        builder.insertCell();
        builder.write("Odd row");
        builder.endTable();

        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.getBorders().setColor(Color.BLACK);
        tableStyle.getBorders().setLineStyle(LineStyle.DOT_DASH);
        // Define our stripe through one column and row
        tableStyle.setColumnStripe(1);
        tableStyle.setRowStripe(1);
        // Let's start from the first row and second column
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.ODD_ROW_BANDING).getShading().setBackgroundPatternColor(msColor.getAliceBlue());
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.EVEN_COLUMN_BANDING).getShading().setBackgroundPatternColor(msColor.getAliceBlue());
        
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        table.setStyle(tableStyle);

        doc.save(getArtifactsDir() + "Table.WorkWithOddEvenRowColumnStyles.docx");
        //ExEnd
    }

    @Test
    public void convertToHorizontallyMergedCells() throws Exception
    {
        //ExStart
        //ExFor:Table.ConvertToHorizontallyMergedCells
        //ExSummary:Shows how to convert cells horizontally merged by width to cells merged by CellFormat.HorizontalMerge.
        Document doc = new Document(getMyDir() + "Table with merged cells.docx");

        // MS Word does not write merge flags anymore, they define merged cells by its width
        // So AW by default define only 5 cells in a row and all of it didn't have horizontal merge flag
        Table table = doc.getFirstSection().getBody().getTables().get(0);
        Row row = table.getRows().get(0);
        Assert.assertEquals(5, row.getCells().getCount());

        // To resolve this inconvenience, we have added new public method to convert cells which are horizontally merged
        // by its width to the cell horizontally merged by flags. Thus now we have 7 cells and some of them have
        // horizontal merge value
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
