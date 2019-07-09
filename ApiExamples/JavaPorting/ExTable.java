// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.ms.System.msConsole;
import com.aspose.words.Row;
import com.aspose.words.Cell;
import com.aspose.ms.System.msString;
import com.aspose.words.SaveFormat;
import org.testng.Assert;
import com.aspose.words.Node;
import com.aspose.words.TableCollection;
import com.aspose.words.Shape;
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
import com.aspose.words.ControlChar;
import com.aspose.words.Paragraph;
import com.aspose.words.DocumentBuilder;
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
        //ExFor:Table
        //ExFor:Row.Cells
        //ExFor:Table.Rows
        //ExFor:Cell
        //ExFor:Row
        //ExFor:RowCollection
        //ExFor:CellCollection
        //ExFor:NodeCollection.IndexOf(Node)
        //ExSummary:Shows how to iterate through all tables in the document and display the content from each cell.
        Document doc = new Document(getMyDir() + "Table.Document.doc");

        // Here we get all tables from the Document node. You can do this for any other composite node
        // which can contain block level nodes. For example you can retrieve tables from header or from a cell
        // containing another table (nested tables).
        NodeCollection tables = doc.getChildNodes(NodeType.TABLE, true);

        // Iterate through all tables in the document
        for (Table table : tables.<Table>OfType() !!Autoporter error: Undefined expression type )
        {
            // Get the index of the table node as contained in the parent node of the table
            int tableIndex = table.getParentNode().getChildNodes().indexOf(table);
            msConsole.writeLine("Start of Table {0}", tableIndex);

            // Iterate through all rows in the table
            for (Row row : table.getRows().<Row>OfType() !!Autoporter error: Undefined expression type )
            {
                int rowIndex = table.getRows().indexOf(row);
                msConsole.writeLine("\tStart of Row {0}", rowIndex);

                // Iterate through all cells in the row
                for (Cell cell : row.getCells().<Cell>OfType() !!Autoporter error: Undefined expression type )
                {
                    int cellIndex = row.getCells().indexOf(cell);
                    // Get the plain text content of this cell.
                    String cellText = msString.trim(cell.toString(SaveFormat.TEXT));
                    // Print the content of the cell.
                    msConsole.writeLine("\t\tContents of Cell:{0} = \"{1}\"", cellIndex, cellText);
                }

                msConsole.writeLine("\tEnd of Row {0}", rowIndex);
            }

            msConsole.writeLine("End of Table {0}", tableIndex);
            msConsole.writeLine();
        }
        //ExEnd

        Assert.That(tables.getCount(), Is.GreaterThan(0));
    }

    @Test
    public void calculateDepthOfNestedTables() throws Exception
    {
        //ExStart
        //ExFor:Node.GetAncestor(NodeType)
        //ExFor:Table.NodeType
        //ExFor:Cell.Tables
        //ExFor:TableCollection
        //ExFor:NodeCollection.Count
        //ExSummary:Shows how to find out if a table contains another table or if the table itself is nested inside another table.
        Document doc = new Document(getMyDir() + "Table.NestedTables.doc");
        int tableIndex = 0;

        for (Table table : doc.getChildNodes(NodeType.TABLE, true).<Table>OfType() !!Autoporter error: Undefined expression type )
        {
            // First lets find if any cells in the table have tables themselves as children.
            int count = getChildTableCount(table);
            msConsole.writeLine("Table #{0} has {1} tables directly within its cells", tableIndex, count);

            // Now let's try the other way around, lets try find if the table is nested inside another table and at what depth.
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

        /*NodeType*/int type = table.getNodeType();
        // The parent of the table will be a Cell, instead attempt to find a grandparent that is of type Table
        Node parent = table.getAncestor(type);

        while (parent != null)
        {
            // Every time we find a table a level up we increase the depth counter and then try to find an
            // ancestor of type table from the parent.
            depth++;
            parent = parent.getAncestor(type);
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
    public void convertTextboxToTable() throws Exception
    {
        //ExStart
        //ExId:TextboxToTable
        //ExSummary:Shows how to convert a textbox to a table and retain almost identical formatting. This is useful for HTML export.
        // Open the document
        Document doc = new Document(getMyDir() + "Shape.TextBox.doc");

        // Convert all shape nodes which contain child nodes.
        // We convert the collection to an array as static "snapshot" because the original textboxes will be removed after conversion which will
        // invalidate the enumerator.
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

        // Create a table to replace the textbox and transfer the same content and formatting.
        Table table = new Table(doc);
        // Ensure that the table contains a row and a cell.
        table.ensureMinimum();
        // Use fixed column widths.
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        // A shape is inline level (within a paragraph) where a table can only be block level so insert the table
        // after the paragraph which contains the shape.
        Node shapeParent = textBox.getParentNode();
        shapeParent.getParentNode().insertAfter(table, shapeParent);

        // If the textbox is not inline then try to match the shape's left position using the table's left indent.
        if (!textBox.isInline() && textBox.getLeft() < section.getPageSetup().getPageWidth())
            table.setLeftIndent(textBox.getLeft());

        // We are only using one cell to replicate a textbox so we can make use of the FirstRow and FirstCell property.
        // Carry over borders and shading.
        Row firstRow = table.getFirstRow();
        Cell firstCell = firstRow.getFirstCell();
        firstCell.getCellFormat().getBorders().setColor(textBox.getStrokeColor());
        firstCell.getCellFormat().getBorders().setLineWidth(textBox.getStrokeWeight());
        firstCell.getCellFormat().getShading().setBackgroundPatternColor(textBox.getFill().getColor());

        // Transfer the same height and width of the textbox to the table.
        firstRow.getRowFormat().setHeightRule(HeightRule.EXACTLY);
        firstRow.getRowFormat().setHeight(textBox.getHeight());
        firstCell.getCellFormat().setWidth(textBox.getWidth());
        table.setAllowAutoFit(false);

        // Replicate the textbox's horizontal alignment.
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
                // Most other options are left by default.
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

        // Remove the empty textbox from the document.
        textBox.remove();
    }
    //ExEnd

    @Test
    public void ensureTableMinimum() throws Exception
    {
        //ExStart
        //ExFor:Table.EnsureMinimum
        //ExSummary:Shows how to ensure a table node is valid.
        Document doc = new Document();

        // Create a new table and add it to the document.
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);

        // Ensure the table is valid (has at least one row with one cell).
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

        // Create a new table and add it to the document.
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);

        // Create a new row and add it to the table.
        Row row = new Row(doc);
        table.appendChild(row);

        // Ensure the row is valid (has at least one cell).
        row.ensureMinimum();
        //ExEnd
    }

    @Test
    public void ensureCellMinimum() throws Exception
    {
        //ExStart
        //ExFor:Cell.EnsureMinimum
        //ExSummary:Shows how to ensure a cell node is valid.
        Document doc = new Document(getMyDir() + "Table.Document.doc");

        // Gets the first cell in the document.
        Cell cell = (Cell) doc.getChild(NodeType.CELL, 0, true);

        // Ensure the cell is valid (the last child is a paragraph).
        cell.ensureMinimum();
        //ExEnd
    }

    @Test
    public void setTableBordersOutline() throws Exception
    {
        //ExStart
        //ExFor:Table.Alignment
        //ExFor:TableAlignment
        //ExFor:Table.ClearBorders
        //ExFor:Table.SetBorder
        //ExFor:TextureIndex
        //ExFor:Table.SetShading
        //ExId:TableBordersOutline
        //ExSummary:Shows how to apply a outline border to a table.
        Document doc = new Document(getMyDir() + "Table.EmptyTable.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Align the table to the center of the page.
        table.setAlignment(TableAlignment.CENTER);

        // Clear any existing borders from the table.
        table.clearBorders();

        // Set a green border around the table but not inside. 
        table.setBorder(BorderType.LEFT, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.RIGHT, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.TOP, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.BOTTOM, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);

        // Fill the cells with a light green solid color.
        table.setShading(TextureIndex.TEXTURE_SOLID, msColor.getLightGreen(), msColor.Empty);

        doc.save(getArtifactsDir() + "Table.SetOutlineBorders.doc");
        //ExEnd

        // Verify the borders were set correctly.
        msAssert.areEqual(TableAlignment.CENTER, table.getAlignment());
        msAssert.areEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getTop().getColor().getRGB());
        msAssert.areEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getLeft().getColor().getRGB());
        msAssert.areEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getRight().getColor().getRGB());
        msAssert.areEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getBottom().getColor().getRGB());
        msAssert.areNotEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getHorizontal().getColor().getRGB());
        msAssert.areNotEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getVertical().getColor().getRGB());
        msAssert.areEqual(msColor.getLightGreen().getRGB(),
            table.getFirstRow().getFirstCell().getCellFormat().getShading().getForegroundPatternColor().getRGB());
    }

    @Test
    public void setTableBordersAll() throws Exception
    {
        //ExStart
        //ExFor:Table.SetBorders
        //ExId:TableBordersAll
        //ExSummary:Shows how to build a table with all borders enabled (grid).
        Document doc = new Document(getMyDir() + "Table.EmptyTable.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Clear any existing borders from the table.
        table.clearBorders();

        // Set a green border around and inside the table.
        table.setBorders(LineStyle.SINGLE, 1.5, msColor.getGreen());

        doc.save(getArtifactsDir() + "Table.SetAllBorders.doc");
        //ExEnd

        // Verify the borders were set correctly.
        msAssert.areEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getTop().getColor().getRGB());
        msAssert.areEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getLeft().getColor().getRGB());
        msAssert.areEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getRight().getColor().getRGB());
        msAssert.areEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getBottom().getColor().getRGB());
        msAssert.areEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getHorizontal().getColor().getRGB());
        msAssert.areEqual(msColor.getGreen().getRGB(), table.getFirstRow().getRowFormat().getBorders().getVertical().getColor().getRGB());
    }

    @Test
    public void rowFormatProperties() throws Exception
    {
        //ExStart
        //ExFor:RowFormat
        //ExFor:Row.RowFormat
        //ExId:RowFormatProperties
        //ExSummary:Shows how to modify formatting of a table row.
        Document doc = new Document(getMyDir() + "Table.Document.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Retrieve the first row in the table.
        Row firstRow = table.getFirstRow();

        // Modify some row level properties.
        firstRow.getRowFormat().getBorders().setLineStyle(LineStyle.NONE);
        firstRow.getRowFormat().setHeightRule(HeightRule.AUTO);
        firstRow.getRowFormat().setAllowBreakAcrossPages(true);
        //ExEnd

        doc.save(getArtifactsDir() + "Table.RowFormat.doc");

        doc = new Document(getArtifactsDir() + "Table.RowFormat.doc");
        table = (Table)doc.getChild(NodeType.TABLE, 0, true);
        msAssert.areEqual(LineStyle.NONE, table.getFirstRow().getRowFormat().getBorders().getLineStyle());
        msAssert.areEqual(HeightRule.AUTO, table.getFirstRow().getRowFormat().getHeightRule());
        Assert.assertTrue(table.getFirstRow().getRowFormat().getAllowBreakAcrossPages());
    }

    @Test
    public void cellFormatProperties() throws Exception
    {
        //ExStart
        //ExFor:CellFormat
        //ExFor:Cell.CellFormat
        //ExId:CellFormatProperties
        //ExSummary:Shows how to modify formatting of a table cell.
        Document doc = new Document(getMyDir() + "Table.Document.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Retrieve the first cell in the table.
        Cell firstCell = table.getFirstRow().getFirstCell();

        // Modify some row level properties.
        firstCell.getCellFormat().setWidth(30.0); // in points
        firstCell.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        firstCell.getCellFormat().getShading().setForegroundPatternColor(msColor.getLightGreen());
        //ExEnd

        doc.save(getArtifactsDir() + "Table.CellFormat.doc");

        doc = new Document(getArtifactsDir() + "Table.CellFormat.doc");
        table = (Table)doc.getChild(NodeType.TABLE, 0, true);
        msAssert.areEqual(30, table.getFirstRow().getFirstCell().getCellFormat().getWidth());
        msAssert.areEqual(TextOrientation.DOWNWARD, table.getFirstRow().getFirstCell().getCellFormat().getOrientation());
        msAssert.areEqual(msColor.getLightGreen().getRGB(),
            table.getFirstRow().getFirstCell().getCellFormat().getShading().getForegroundPatternColor().getRGB());
    }

    @Test
    public void getDistance() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table.Distance.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        msAssert.areEqual(11.35d, table.getDistanceTop());
        msAssert.areEqual(26.35d, table.getDistanceBottom());
        msAssert.areEqual(9.05d, table.getDistanceLeft());
        msAssert.areEqual(22.7d, table.getDistanceRight());
    }

    @Test
    public void removeBordersFromAllCells() throws Exception
    {
        //ExStart
        //ExFor:Table
        //ExFor:Table.ClearBorders
        //ExSummary:Shows how to remove all borders from a table.
        Document doc = new Document(getMyDir() + "Table.Document.doc");

        // Remove all borders from the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Clear the borders all cells in the table.
        table.clearBorders();

        doc.save(getArtifactsDir() + "Table.ClearBorders.doc");
        //ExEnd
    }

    @Test
    public void replaceTextInTable() throws Exception
    {
        //ExStart
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExFor:Cell
        //ExId:ReplaceTextTable
        //ExSummary:Shows how to replace all instances of String of text in a table and cell.
        Document doc = new Document(getMyDir() + "Table.SimpleTable.doc");

        // Get the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(true);
        options.setFindWholeWordsOnly(true);

        // Replace any instances of our String in the entire table.
        table.getRange().replace("Carrots", "Eggs", options);
        // Replace any instances of our String in the last cell of the table only.
        table.getLastRow().getLastCell().getRange().replace("50", "20", options);

        doc.save(getArtifactsDir() + "Table.ReplaceCellText.doc");
        //ExEnd

        msAssert.areEqual("20", msString.trim(table.getLastRow().getLastCell().toString(SaveFormat.TEXT)));
    }

    @Test
    public void printTableRange() throws Exception
    {
        //ExStart
        //ExId:PrintTableRange
        //ExSummary:Shows how to print the text range of a table.
        Document doc = new Document(getMyDir() + "Table.SimpleTable.doc");

        // Get the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // The range text will include control characters such as "\a" for a cell.
        // You can call ToString on the desired node to retrieve the plain text content.

        // Print the plain text range of the table to the screen.
        msConsole.writeLine("Contents of the table: ");
        msConsole.writeLine(table.getRange().getText());
        //ExEnd

        //ExStart
        //ExId:PrintRowAndCellRange
        //ExSummary:Shows how to print the text range of row and table elements.
        // Print the contents of the second row to the screen.
        msConsole.writeLine("\nContents of the row: ");
        msConsole.writeLine(table.getRows().get(1).getRange().getText());

        // Print the contents of the last cell in the table to the screen.
        msConsole.writeLine("\nContents of the cell: ");
        msConsole.writeLine(table.getLastRow().getLastCell().getRange().getText());
        //ExEnd

        msAssert.areEqual("Apples\r" + ControlChar.CELL + "20\r" + ControlChar.CELL + ControlChar.CELL,
            table.getRows().get(1).getRange().getText());
        msAssert.areEqual("50\r\u0007", table.getLastRow().getLastCell().getRange().getText());
    }

    @Test
    public void cloneTable() throws Exception
    {
        //ExStart
        //ExId:CloneTable
        //ExSummary:Shows how to make a clone of a table in the document and insert it after the original table.
        Document doc = new Document(getMyDir() + "Table.SimpleTable.doc");

        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Create a clone of the table.
        Table tableClone = (Table) table.deepClone(true);

        // Insert the cloned table into the document after the original
        table.getParentNode().insertAfter(tableClone, table);

        // Insert an empty paragraph between the two tables or else they will be combined into one
        // upon save. This has to do with document validation.
        table.getParentNode().insertAfter(new Paragraph(doc), table);

        doc.save(getArtifactsDir() + "Table.CloneTableAndInsert.doc");
        //ExEnd

        // Verify that the table was cloned and inserted properly.
        msAssert.areEqual(2, doc.getChildNodes(NodeType.TABLE, true).getCount());
        msAssert.areEqual(table.getRange().getText(), tableClone.getRange().getText());

        //ExStart
        //ExId:CloneTableRemoveContent
        //ExSummary:Shows how to remove all content from the cells of a cloned table.
        for (Cell cell : tableClone.getChildNodes(NodeType.CELL, true).<Cell>OfType() !!Autoporter error: Undefined expression type )
            cell.removeAllChildren();
        //ExEnd

        msAssert.areEqual("", msString.trim(tableClone.toString(SaveFormat.TEXT)));
    }

    @Test
    public void rowFormatDisableBreakAcrossPages() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table.TableAcrossPage.doc");

        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        //ExStart
        //ExFor:RowFormat.AllowBreakAcrossPages
        //ExId:RowFormatAllowBreaks
        //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
        // Disable breaking across pages for all rows in the table.
        for (Row row : table.<Row>OfType() !!Autoporter error: Undefined expression type )
            row.getRowFormat().setAllowBreakAcrossPages(false);
        //ExEnd

        doc.save(getArtifactsDir() + "Table.DisableBreakAcrossPages.doc");

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
        //ExId:AllowAutoFit
        //ExSummary:Shows how to set a table to shrink or grow each cell to accommodate its contents.
        table.setAllowAutoFit(true);
        //ExEnd
    }

    @Test
    public void keepTableTogether() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table.TableAcrossPage.doc");

        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        //ExStart
        //ExFor:ParagraphFormat.KeepWithNext
        //ExFor:Row.IsLastRow
        //ExFor:Paragraph.IsEndOfCell
        //ExFor:Paragraph.IsInCell
        //ExFor:Cell.ParentRow
        //ExFor:Cell.Paragraphs
        //ExId:KeepTableTogether
        //ExSummary:Shows how to set a table to stay together on the same page.
        // To keep a table from breaking across a page we need to enable KeepWithNext 
        // for every paragraph in the table except for the last paragraphs in the last 
        // row of the table.
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

        // Verify the correct paragraphs were set properly.
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
        //ExId:AddClonedRowToTable
        //ExSummary:Shows how to make a clone of the last row of a table and append it to the table.
        Document doc = new Document(getMyDir() + "Table.SimpleTable.doc");

        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Clone the last row in the table.
        Row clonedRow = (Row) table.getLastRow().deepClone(true);

        // Remove all content from the cloned row's cells. This makes the row ready for
        // new content to be inserted into.
        for (Cell cell : clonedRow.getCells().<Cell>OfType() !!Autoporter error: Undefined expression type )
            cell.removeAllChildren();

        // Add the row to the end of the table.
        table.appendChild(clonedRow);

        doc.save(getArtifactsDir() + "Table.AddCloneRowToTable.doc");
        //ExEnd

        // Verify that the row was cloned and appended properly.
        msAssert.areEqual(5, table.getRows().getCount());
        msAssert.areEqual("", msString.trim(table.getLastRow().toString(SaveFormat.TEXT)));
        msAssert.areEqual(2, table.getLastRow().getCells().getCount());
    }

    @Test
    public void fixDefaultTableWidthsInAw105() throws Exception
    {
        //ExStart
        //ExId:FixTablesDefaultFixedColumnWidth
        //ExSummary:Shows how to revert the default behavior of table sizing to use column widths.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Keep a reference to the table being built.
        Table table = builder.startTable();

        // Apply some formatting.
        builder.getCellFormat().setWidth(100.0);
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.RED);

        builder.insertCell();
        // This will cause the table to be structured using column widths as in previous versions
        // instead of fitted to the page width like in the newer versions.
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        // Continue with building your table as usual...
        //ExEnd
    }

    @Test
    public void fixDefaultTableBordersIn105() throws Exception
    {
        //ExStart
        //ExId:FixTablesDefaultBorders
        //ExSummary:Shows how to revert the default borders on tables back to no border lines.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Keep a reference to the table being built.
        Table table = builder.startTable();

        builder.insertCell();
        // Clear all borders to match the defaults used in previous versions.
        table.clearBorders();

        // Continue with building your table as usual...
        //ExEnd
    }

    @Test
    public void fixDefaultTableFormattingExceptionIn105() throws Exception
    {
        //ExStart
        //ExId:FixTableFormattingException
        //ExSummary:Shows how to avoid encountering an exception when applying table formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Keep a reference to the table being built.
        Table table = builder.startTable();

        // We must first insert a new cell which in turn inserts a row into the table.
        builder.insertCell();
        // Once a row exists in our table we can apply table wide formatting.
        table.setAllowAutoFit(true);

        // Continue with building your table as usual...
        //ExEnd
    }

    @Test
    public void fixRowFormattingNotAppliedIn105() throws Exception
    {
        //ExStart
        //ExId:FixRowFormattingNotApplied
        //ExSummary:Shows how to fix row formatting not being applied to some rows.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();

        // For the first row this will be set correctly.
        builder.getRowFormat().setHeadingFormat(true);

        builder.insertCell();
        builder.writeln("Text");
        builder.insertCell();
        builder.writeln("Text");

        // End the first row.
        builder.endRow();

        // Here we would normally define some other row formatting, such as disabling the 
        // heading format. However at the moment this will be ignored and the value from the 
        // first row reapplied to the row.

        builder.insertCell();

        // Instead make sure to specify the row formatting for the second row here.
        builder.getRowFormat().setHeadingFormat(false);

        // Continue with building your table as usual...
        //ExEnd
    }

    @Test
    public void getIndexOfTableElements() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table.Document.doc");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        //ExStart
        //ExFor:NodeCollection.IndexOf
        //ExId:IndexOfTable
        //ExSummary:Retrieves the index of a table in the document.
        NodeCollection allTables = doc.getChildNodes(NodeType.TABLE, true);
        int tableIndex = allTables.indexOf(table);
        //ExEnd

        Row row = table.getRows().get(2);
        //ExStart
        //ExFor:Row
        //ExFor:CompositeNode.IndexOf
        //ExId:IndexOfRow
        //ExSummary:Retrieves the index of a row in a table.
        int rowIndex = table.indexOf(row);
        //ExEnd

        Cell cell = row.getLastCell();
        //ExStart
        //ExFor:Cell
        //ExFor:CompositeNode.IndexOf
        //ExId:IndexOfCell
        //ExSummary:Retrieves the index of a cell in a row.
        int cellIndex = row.indexOf(cell);
        //ExEnd

        msAssert.areEqual(0, tableIndex);
        msAssert.areEqual(2, rowIndex);
        msAssert.areEqual(4, cellIndex);
    }

    @Test
    public void getPreferredWidthTypeAndValue() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table.Document.doc");

        // Find the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        //ExStart
        //ExFor:PreferredWidthType
        //ExFor:PreferredWidth.Type
        //ExFor:PreferredWidth.Value
        //ExId:GetPreferredWidthTypeAndValue
        //ExSummary:Retrieves the preferred width type of a table cell.
        Cell firstCell = table.getFirstRow().getFirstCell();
        /*PreferredWidthType*/int type = firstCell.getCellFormat().getPreferredWidth().getType();
        double value = firstCell.getCellFormat().getPreferredWidth().getValue();
        //ExEnd

        msAssert.areEqual(PreferredWidthType.PERCENT, type);
        msAssert.areEqual(11.16, value);
    }

    @Test
    public void insertTableUsingNodeConstructors() throws Exception
    {
        //ExStart
        //ExFor:Table
        //ExFor:Table.AllowCellSpacing
        //ExFor:Row
        //ExFor:Row.RowFormat
        //ExFor:RowFormat
        //ExFor:Cell
        //ExFor:Cell.CellFormat
        //ExFor:CellFormat
        //ExFor:CellFormat.Shading
        //ExFor:Cell.FirstParagraph
        //ExId:InsertTableUsingNodeConstructors
        //ExSummary:Shows how to insert a table using the constructors of nodes.
        Document doc = new Document();

        // We start by creating the table object. Note how we must pass the document object
        // to the constructor of each node. This is because every node we create must belong
        // to some document.
        Table table = new Table(doc);
        // Add the table to the document.
        doc.getFirstSection().getBody().appendChild(table);

        // Here we could call EnsureMinimum to create the rows and cells for us. This method is used
        // to ensure that the specified node is valid, in this case a valid table should have at least one
        // row and one cell, therefore this method creates them for us.

        // Instead we will handle creating the row and table ourselves. This would be the best way to do this
        // if we were creating a table inside an algorithm for example.
        Row row = new Row(doc);
        row.getRowFormat().setAllowBreakAcrossPages(true);
        table.appendChild(row);

        // We can now apply any auto fit settings.
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        // Create a cell and add it to the row
        Cell cell = new Cell(doc);
        cell.getCellFormat().getShading().setBackgroundPatternColor(Color.LightBlue);
        cell.getCellFormat().setWidth(80.0);

        // Add a paragraph to the cell as well as a new run with some text.
        cell.appendChild(new Paragraph(doc));
        cell.getFirstParagraph().appendChild(new Run(doc, "Row 1, Cell 1 Text"));

        // Add the cell to the row.
        row.appendChild(cell);

        // We would then repeat the process for the other cells and rows in the table.
        // We can also speed things up by cloning existing cells and rows.
        row.appendChild(cell.deepClone(false));
        row.getLastCell().appendChild(new Paragraph(doc));
        row.getLastCell().getFirstParagraph().appendChild(new Run(doc, "Row 1, Cell 2 Text"));

        // Remove spacing between cells
        table.setAllowCellSpacing(false);

        doc.save(getArtifactsDir() + "Table.InsertTableUsingNodes.doc");
        //ExEnd

        msAssert.areEqual(1, doc.getChildNodes(NodeType.TABLE, true).getCount());
        msAssert.areEqual(1, doc.getChildNodes(NodeType.ROW, true).getCount());
        msAssert.areEqual(2, doc.getChildNodes(NodeType.CELL, true).getCount());
        msAssert.areEqual("Row 1, Cell 1 Text\r\nRow 1, Cell 2 Text",
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
    //ExId:NestedTableNodeConstructors
    //ExSummary:Shows how to build a nested table without using DocumentBuilder.
    @Test //ExSkip
    public void nestedTablesUsingNodeConstructors() throws Exception
    {
        Document doc = new Document();

        // Create the outer table with three rows and four columns.
        Table outerTable = createTable(doc, 3, 4, "Outer Table");
        // Add it to the document body.
        doc.getFirstSection().getBody().appendChild(outerTable);

        // Create another table with two rows and two columns.
        Table innerTable = createTable(doc, 2, 2, "Inner Table");
        // Add this table to the first cell of the outer table.
        outerTable.getFirstRow().getFirstCell().appendChild(innerTable);

        doc.save(getArtifactsDir() + "Table.CreateNestedTable.doc");

        msAssert.areEqual(2, doc.getChildNodes(NodeType.TABLE, true).getCount()); // ExSkip
        msAssert.areEqual(1, outerTable.getFirstRow().getFirstCell().getTables().getCount()); //ExSkip
        msAssert.areEqual(16, outerTable.getChildNodes(NodeType.CELL, true).getCount()); //ExSkip
        msAssert.areEqual(4, innerTable.getChildNodes(NodeType.CELL, true).getCount()); //ExSkip
        msAssert.areEqual("Aspose table title", innerTable.getTitle()); //ExSkip
        msAssert.areEqual("Aspose table description", innerTable.getDescription()); //ExSkip
    }

    /// <summary>
    /// Creates a new table in the document with the given dimensions and text in each cell.
    /// </summary>
    private Table createTable(Document doc, int rowCount, int cellCount, String cellText) throws Exception
    {
        Table table = new Table(doc);

        // Create the specified number of rows.
        for (int rowId = 1; rowId <= rowCount; rowId++)
        {
            Row row = new Row(doc);
            table.appendChild(row);

            // Create the specified number of cells for each row.
            for (int cellId = 1; cellId <= cellCount; cellId++)
            {
                Cell cell = new Cell(doc);
                row.appendChild(cell);
                // Add a blank paragraph to the cell.
                cell.appendChild(new Paragraph(doc));

                // Add the text.
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
    //ExId:CheckCellMerge
    //ExSummary:Prints the horizontal and vertical merge type of a cell.
    @Test //ExSkip
    public void checkCellsMerged() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table.MergedCells.doc");

        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        for (Row row : table.getRows().<Row>OfType() !!Autoporter error: Undefined expression type )
        {
            for (Cell cell : row.getCells().<Cell>OfType() !!Autoporter error: Undefined expression type )
            {
                msConsole.writeLine(printCellMergeType(cell));
            }
        }

        msAssert.areEqual("The cell at R1, C1 is horizontally merged.",
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
        else if (isHorizontallyMerged)
            return $"The cell at {cellLocation} is horizontally merged.";
        else if (isVerticallyMerged)
            return $"The cell at {cellLocation} is vertically merged";
        else
            return $"The cell at {cellLocation} is not merged";
    }
    //ExEnd

    @Test
    public void mergeCellRange() throws Exception
    {
        // Open the document
        Document doc = new Document(getMyDir() + "Table.Document.doc");

        // Retrieve the first table in the body of the first section.
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        //ExStart
        //ExId:MergeCellRange
        //ExSummary:Merges the range of cells between the two specified cells.
        // We want to merge the range of cells found in between these two cells.
        Cell cellStartRange = table.getRows().get(2).getCells().get(2);
        Cell cellEndRange = table.getRows().get(3).getCells().get(3);

        // Merge all the cells between the two specified cells into one.
        mergeCells(cellStartRange, cellEndRange);
        //ExEnd

        // Save the document.
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

        msAssert.areEqual(4, mergedCellsCount);
        Assert.assertTrue(table.getRows().get(2).getCells().get(2).getCellFormat().getHorizontalMerge() == CellMerge.FIRST);
        Assert.assertTrue(table.getRows().get(2).getCells().get(2).getCellFormat().getVerticalMerge() == CellMerge.FIRST);
        Assert.assertTrue(table.getRows().get(3).getCells().get(3).getCellFormat().getHorizontalMerge() == CellMerge.PREVIOUS);
        Assert.assertTrue(table.getRows().get(3).getCells().get(3).getCellFormat().getVerticalMerge() == CellMerge.PREVIOUS);
    }

    //ExStart
    //ExId:MergeCellsMethod
    //ExSummary:A method which merges all cells of a table in the specified range of cells.
    /// <summary>
    /// Merges the range of cells found between the two specified cells both horizontally and vertically. Can span over multiple rows.
    /// </summary>
    @Test (enabled = false)
    public static void mergeCells(Cell startCell, Cell endCell)
    {
        Table parentTable = startCell.getParentRow().getParentTable();

        // Find the row and cell indices for the start and end cell.
        /*Point*/long startCellPos = msPoint.ctor(startCell.getParentRow().indexOf(startCell),
            parentTable.indexOf(startCell.getParentRow()));
        /*Point*/long endCellPos = msPoint.ctor(endCell.getParentRow().indexOf(endCell), parentTable.indexOf(endCell.getParentRow()));
        // Create the range of cells to be merged based off these indices. Inverse each index if the end cell if before the start cell. 
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
                // Check if the current cell is inside our merge range then merge it.
                if (mergeRange.contains(currentPos))
                {
                    cell.getCellFormat().setHorizontalMerge(msPoint.getX(currentPos) == mergeRange.getX() ? CellMerge.FIRST : CellMerge.PREVIOUS);
                    cell.getCellFormat().setVerticalMerge(msPoint.getY(currentPos) == mergeRange.getY() ? CellMerge.FIRST : CellMerge.PREVIOUS);
                }
            }
        }
    }
    //ExEnd

    @Test
    public void combineTables() throws Exception
    {
        //ExStart
        //ExFor:Table
        //ExFor:Cell.CellFormat
        //ExFor:CellFormat.Borders
        //ExFor:Table.Rows
        //ExFor:Table.FirstRow
        //ExFor:CellFormat.ClearFormatting
        //ExId:CombineTables
        //ExSummary:Shows how to combine the rows from two tables into one.
        // Load the document.
        Document doc = new Document(getMyDir() + "Table.Document.doc");

        // Get the first and second table in the document.
        // The rows from the second table will be appended to the end of the first table.
        Table firstTable = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Table secondTable = (Table) doc.getChild(NodeType.TABLE, 1, true);

        // Append all rows from the current table to the next.
        // Due to the design of tables even tables with different cell count and widths can be joined into one table.
        while (secondTable.hasChildNodes())
            firstTable.getRows().add(secondTable.getFirstRow());

        // Remove the empty table container.
        secondTable.remove();

        doc.save(getArtifactsDir() + "Table.CombineTables.doc");
        //ExEnd

        msAssert.areEqual(1, doc.getChildNodes(NodeType.TABLE, true).getCount());
        msAssert.areEqual(9, doc.getFirstSection().getBody().getTables().get(0).getRows().getCount());
        msAssert.areEqual(42, doc.getFirstSection().getBody().getTables().get(0).getChildNodes(NodeType.CELL, true).getCount());
    }

    @Test
    public void splitTable() throws Exception
    {
        //ExStart
        //ExId:SplitTableAtRow
        //ExSummary:Shows how to split a table into two tables a specific row.
        // Load the document.
        Document doc = new Document(getMyDir() + "Table.SimpleTable.doc");

        // Get the first table in the document.
        Table firstTable = (Table) doc.getChild(NodeType.TABLE, 0, true);

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

        doc.save(getArtifactsDir() + "Table.SplitTable.doc");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.SplitTable.doc");
        // Test we are adding the rows in the correct order and the 
        // selected row was also moved.
        msAssert.areEqual(row, table.getFirstRow());

        msAssert.areEqual(2, firstTable.getRows().getCount());
        msAssert.areEqual(2, table.getRows().getCount());
        msAssert.areEqual(2, doc.getChildNodes(NodeType.TABLE, true).getCount());
    }

    @Test
    public void checkDefaultValuesForFloatingTableProperties() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = DocumentHelper.insertTable(builder);

        if (table.getTextWrapping() == TextWrapping.AROUND)
        {
            msAssert.areEqual(HorizontalAlignment.DEFAULT, table.getRelativeHorizontalAlignment());
            msAssert.areEqual(VerticalAlignment.DEFAULT, table.getRelativeVerticalAlignment());
            msAssert.areEqual(RelativeHorizontalPosition.COLUMN, table.getHorizontalAnchor());
            msAssert.areEqual(RelativeVerticalPosition.MARGIN, table.getVerticalAnchor());
            msAssert.areEqual(0, table.getAbsoluteHorizontalDistance());
            msAssert.areEqual(0, table.getAbsoluteVerticalDistance());
            msAssert.areEqual(true, table.getAllowOverlap());
        }
    }

    @Test
    public void floatingTableProperties() throws Exception
    {
        //ExStart
        //ExFor:Table.RelativeHorizontalAlignment
        //ExFor:Table.RelativeVerticalAlignment
        //ExFor:Table.HorizontalAnchor
        //ExFor:Table.VerticalAnchor
        //ExFor:Table.AbsoluteHorizontalDistance
        //ExFor:Table.AbsoluteVerticalDistance
        //ExFor:Table.AllowOverlap
        //ExFor:ShapeBase.AllowOverlap
        //ExSummary:Shows how get properties for floating tables
        Document doc = new Document(getMyDir() + "Table.Distance.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        if (table.getTextWrapping() == TextWrapping.AROUND)
        {
            msAssert.areEqual(HorizontalAlignment.DEFAULT, table.getRelativeHorizontalAlignment());
            msAssert.areEqual(VerticalAlignment.DEFAULT, table.getRelativeVerticalAlignment());
            msAssert.areEqual(RelativeHorizontalPosition.MARGIN, table.getHorizontalAnchor());
            msAssert.areEqual(RelativeVerticalPosition.PARAGRAPH, table.getVerticalAnchor());
            msAssert.areEqual(0, table.getAbsoluteHorizontalDistance());
            msAssert.areEqual(4.8, table.getAbsoluteVerticalDistance());
            msAssert.areEqual(true, table.getAllowOverlap());
        }
        //ExEnd
    }

    @Test
    public void tableStyleCreation() throws Exception
    {
        //ExStart
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
        //ExSummary:Shows how to work with certain area styles of a table.
        Document doc = new Document(getMyDir() + "Table.ConditionalStyles.docx");

        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        // There is a different ways how to get conditional styles:
        // by conditional style type
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.FIRST_ROW).getShading().setBackgroundPatternColor(msColor.getAliceBlue());
        // by index
        tableStyle.getConditionalStyles().get(0).getBorders().setColor(Color.BLACK);
        tableStyle.getConditionalStyles().get(0).getBorders().setLineStyle(LineStyle.DOT_DASH);
        msAssert.areEqual(ConditionalStyleType.FIRST_ROW, tableStyle.getConditionalStyles().get(0).getType());
        // directly from ConditionalStyleCollection
        tableStyle.getConditionalStyles().getFirstRow().getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        // To see this in Word document select Total Row checkbox in Design Tab
        tableStyle.getConditionalStyles().getLastRow().setBottomPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setLeftPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setRightPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setTopPadding(10.0);
        // To see this in Word document select Last Column checkbox in Design Tab
        tableStyle.getConditionalStyles().getLastColumn().getFont().setBold(true);

        msConsole.writeLine(tableStyle.getConditionalStyles().getCount());
        msConsole.writeLine(tableStyle.getConditionalStyles().get(0).getType());

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
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
        Document doc = new Document(getMyDir() + "Table.ConditionalStyles.docx");

        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        // You can reset styles from the specific table area
        tableStyle.getConditionalStyles().get(0).clearFormatting();
        // Or clear all table styles
        tableStyle.getConditionalStyles().clearFormatting();
        //ExEnd
    }

    @Test (enabled = false, description = "WORDSNET-18708")
    public void getConditionalStylesEnumerator() throws Exception
    {
        //ExStart
        //ExFor:ConditionalStyle.Type
        //ExFor:ConditionalStyleCollection.GetEnumerator
        //ExSummary:Shows how to enumerate all table styles in a collection.
        Document doc = new Document(getMyDir() + "Table.ConditionalStyles.docx");

        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");

        // Get the enumerator from the document's ConditionalStyleCollection and iterate over the styles
        Iterator<ConditionalStyle> enumerator = tableStyle.getConditionalStyles().iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                ConditionalStyle currentStyle = enumerator.next();

                if (currentStyle != null)
                {
                    msConsole.writeLine(currentStyle.getType());
                }
            }
        }
        finally { if (enumerator != null) enumerator.close(); }
        //ExEnd
    }

    @Test
    public void workWithOddEvenRowColumnStyles() throws Exception
    {
        //ExStart
        //ExFor:TableStyle.ColumnStripe
        //ExFor:TableStyle.RowStripe
        //ExSummary:Shows how to work with odd/even row/column styles.
        Document doc = new Document(getMyDir() + "Table.ConditionalStyles.docx");

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
        Document doc = new Document(getMyDir() + "Table.ConvertToHorizontallyMergedCells.docx");

        // MS Word does not write merge flags anymore, they define merged cells by its width
        // So AW by default define only 5 cells in a row and all of it didn't have horizontal merge flag
        Table table = doc.getFirstSection().getBody().getTables().get(0);
        Row row = table.getRows().get(0);
        msAssert.areEqual(5, row.getCells().getCount());

        // To resolve this inconvenience, we have added new public method to convert cells which are horizontally merged
        // by its width to the cell horizontally merged by flags. Thus now we have 7 cells and some of them have
        // horizontal merge value
        table.convertToHorizontallyMergedCells();
        row = table.getRows().get(0);
        msAssert.areEqual(7, row.getCells().getCount());

        msAssert.areEqual(CellMerge.NONE, row.getCells().get(0).getCellFormat().getHorizontalMerge());
        msAssert.areEqual(CellMerge.FIRST, row.getCells().get(1).getCellFormat().getHorizontalMerge());
        msAssert.areEqual(CellMerge.PREVIOUS, row.getCells().get(2).getCellFormat().getHorizontalMerge());
        msAssert.areEqual(CellMerge.NONE, row.getCells().get(3).getCellFormat().getHorizontalMerge());
        msAssert.areEqual(CellMerge.FIRST, row.getCells().get(4).getCellFormat().getHorizontalMerge());
        msAssert.areEqual(CellMerge.PREVIOUS, row.getCells().get(5).getCellFormat().getHorizontalMerge());
        msAssert.areEqual(CellMerge.NONE, row.getCells().get(6).getCellFormat().getHorizontalMerge());
        //ExEnd
    }
}
