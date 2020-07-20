package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Shape;
import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.awt.*;
import java.text.MessageFormat;
import java.util.Iterator;

/// <summary>
/// Examples using tables in documents.
/// </summary>
@Test
public class ExTable extends ApiExampleBase {
    @Test
    public void createTable() throws Exception {
        //ExStart
        //ExFor:Table
        //ExFor:Row
        //ExFor:Cell
        //ExFor:Table.#ctor(DocumentBase)
        //ExSummary:Shows how to create a simple table.
        Document doc = new Document();

        // Tables are placed in the body of a document
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);

        // Tables contain rows, which contain cells,
        // which contain contents such as paragraphs, runs and even other tables
        // Calling table.EnsureMinimum will also make sure that a table has at least one row, cell and paragraph
        Row firstRow = new Row(doc);
        table.appendChild(firstRow);

        Cell firstCell = new Cell(doc);
        firstRow.appendChild(firstCell);

        Paragraph paragraph = new Paragraph(doc);
        firstCell.appendChild(paragraph);

        Run run = new Run(doc, "Hello world!");
        paragraph.appendChild(run);

        doc.save(getArtifactsDir() + "Table.CreateTable.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.CreateTable.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(1, table.getRows().getCount());
        Assert.assertEquals(1, table.getFirstRow().getCells().getCount());
        Assert.assertEquals("Hello world!", table.getText().trim());
    }

    @Test
    public void rowCellFormat() throws Exception {
        //ExStart
        //ExFor:Row.RowFormat
        //ExFor:RowFormat
        //ExFor:Cell.CellFormat
        //ExFor:CellFormat
        //ExFor:CellFormat.Shading
        //ExSummary:Shows how to modify the format of rows and cells.
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

        // The appearance of rows and individual cells can be edited using the respective formatting objects
        RowFormat rowFormat = table.getFirstRow().getRowFormat();
        rowFormat.setHeight(25.0);
        rowFormat.getBorders().getByBorderType(BorderType.BOTTOM).setColor(Color.RED);

        CellFormat cellFormat = table.getLastRow().getFirstCell().getCellFormat();
        cellFormat.setWidth(100.0);
        cellFormat.getShading().setBackgroundPatternColor(Color.ORANGE);

        doc.save(getArtifactsDir() + "Table.RowCellFormat.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.RowCellFormat.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals("City\u0007Country\u0007\u0007London\u0007U.K.", table.getText().trim());

        rowFormat = table.getFirstRow().getRowFormat();

        Assert.assertEquals(25.0d, rowFormat.getHeight());
        Assert.assertEquals(Color.RED.getRGB(), rowFormat.getBorders().getByBorderType(BorderType.BOTTOM).getColor().getRGB());

        cellFormat = table.getLastRow().getFirstCell().getCellFormat();

        Assert.assertEquals(110.8d, cellFormat.getWidth());
        Assert.assertEquals(Color.ORANGE.getRGB(), cellFormat.getShading().getBackgroundPatternColor().getRGB());
    }

    @Test
    public void displayContentOfTables() throws Exception {
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
        Assert.assertEquals(tables.toArray().length, 2);

        // Iterate through all tables in the document
        for (int i = 0; i < tables.getCount(); i++) {
            // Get the index of the table node as contained in the parent node of the table
            System.out.println(MessageFormat.format("Start of Table {0}", i));

            RowCollection rows = tables.get(i).getRows();

            // Iterate through all rows in the table
            for (int j = 0; j < rows.getCount(); j++) {
                System.out.println(MessageFormat.format("\tStart of Row {0}", j));

                CellCollection cells = rows.get(j).getCells();

                // Iterate through all cells in the row
                for (int k = 0; k < cells.getCount(); k++) {
                    // Get the plain text content of this cell.
                    String cellText = cells.get(k).toString(SaveFormat.TEXT).trim();
                    // Print the content of the cell.
                    System.out.println(MessageFormat.format("\t\tContents of Cell:{0} = \"{1}\"", k, cellText));
                }

                System.out.println(MessageFormat.format("\tEnd of Row {0}", j));
            }

            System.out.println(MessageFormat.format("End of Table {0}\n", i));
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
    //ExSummary:Shows how to find out if a table contains another table or if the table itself is nested inside another table.
    @Test //ExSkip
    public void calculateDepthOfNestedTables() throws Exception {
        Document doc = new Document(getMyDir() + "Nested tables.docx");
        NodeCollection tables = doc.getChildNodes(NodeType.TABLE, true);
        Assert.assertEquals(5, tables.getCount()); //ExSkip

        for (int i = 0; i < tables.getCount(); i++) {
            // First lets find if any cells in the table have tables themselves as children
            int count = getChildTableCount((Table) tables.get(i));
            System.out.println(MessageFormat.format("Table #{0} has {1} tables directly within its cells", i, count));

            // Now let's try the other way around, lets try find if the table is nested inside another table and at what depth
            int tableDepth = getNestedDepthOfTable((Table) tables.get(i));

            if (tableDepth > 0)
                System.out.println(MessageFormat.format("Table #{0} is nested inside another table at depth of {1}", i, tableDepth));
            else
                System.out.println(MessageFormat.format("Table #{0} is a non nested table (is not a child of another table)", i));
        }
    }

    /**
     * Calculates what level a table is nested inside other tables.
     *
     * @returns An integer containing the level the table is nested at.
     * 0 = Table is not nested inside any other table
     * 1 = Table is nested within one parent table
     * 2 = Table is nested within two parent tables etc..
     */
    private static int getNestedDepthOfTable(final Table table) {
        int depth = 0;

        int type = table.getNodeType();
        // The parent of the table will be a Cell, instead attempt to find a grandparent that is of type Table
        Node parent = table.getAncestor(table.getNodeType());

        while (parent != null) {
            // Every time we find a table a level up we increase the depth counter and then try to find an
            // ancestor of type table from the parent
            depth++;
            parent = parent.getAncestor(Table.class);
        }

        return depth;
    }

    /**
     * Determines if a table contains any immediate child table within its cells.
     * Does not recursively traverse through those tables to check for further tables.
     *
     * @returns Returns true if at least one child cell contains a table.
     * Returns false if no cells in the table contains a table.
     */
    private static int getChildTableCount(final Table table) {
        int tableCount = 0;
        // Iterate through all child rows in the table
        for (Row row : table.getRows()) {
            // Iterate through all child cells in the row
            for (Cell cell : row.getCells()) {
                // Retrieve the collection of child tables of this cell
                TableCollection childTables = cell.getTables();

                // If this cell has a table as a child then return true
                if (childTables.getCount() > 0) tableCount++;
            }
        }

        // No cell contains a table
        return tableCount;
    }
    //ExEnd

    @Test
    public void convertTextBoxToTable() throws Exception {
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
        Node[] nodes = doc.getChildNodes(NodeType.SHAPE, true).toArray();
        for (Node node : nodes) {
            Shape shape = (Shape) node;
            if (shape.hasChildNodes()) {
                convertTextboxToTable(shape);
            }
        }

        doc.save(getArtifactsDir() + "Table.ConvertTextBoxToTable.html");
    }

    /**
     * Converts a textbox to a table by copying the same content and formatting.
     * Currently export to HTML will render the textbox as an image which looses any text functionality.
     * This is useful to convert textboxes in order to retain proper text.
     *
     * @param textBox The textbox shape to convert to a table.
     */
    private static void convertTextboxToTable(final Shape textBox) throws Exception {
        if (textBox.getStoryType() != StoryType.TEXTBOX) {
            throw new IllegalArgumentException("Can only convert a shape of type textbox");
        }

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
        if (!textBox.isInline() && textBox.getLeft() < section.getPageSetup().getPageWidth()) {
            table.setLeftIndent(textBox.getLeft());
        }

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
        int horizontalAlignment;
        switch (textBox.getHorizontalAlignment()) {
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
        for (Node node : textBox.getChildNodes(NodeType.ANY, false).toArray()) {
            table.getFirstRow().getFirstCell().appendChild(node);
        }

        // Remove the empty textbox from the document
        textBox.remove();
    }

    @Test
    public void ensureTableMinimum() throws Exception {
        //ExStart
        //ExFor:Table.EnsureMinimum
        //ExSummary:Shows how to ensure a table node is valid.
        Document doc = new Document();

        // Create a new table and add it to the document
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);

        // Currently, the table does not contain any rows, cells or nodes that can have content added to them
        Assert.assertEquals(0, table.getChildNodes(NodeType.ANY, true).getCount());

        // This method ensures that the table has one row, one cell and one paragraph; the minimal nodes required to begin editing
        table.ensureMinimum();
        table.getFirstRow().getFirstCell().getFirstParagraph().appendChild(new Run(doc, "Hello world!"));
        //ExEnd

        Assert.assertEquals(4, table.getChildNodes(NodeType.ANY, true).getCount());
    }

    @Test
    public void ensureRowMinimum() throws Exception {
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

        // Currently, the row does not contain any cells or nodes that can have content added to them
        Assert.assertEquals(0, row.getChildNodes(NodeType.ANY, true).getCount());

        // Ensure the row has at least one cell with one paragraph that we can edit
        row.ensureMinimum();
        row.getFirstCell().getFirstParagraph().appendChild(new Run(doc, "Hello world!"));
        //ExEnd

        Assert.assertEquals(3, row.getChildNodes(NodeType.ANY, true).getCount());
    }

    @Test
    public void ensureCellMinimum() throws Exception {
        //ExStart
        //ExFor:Cell.EnsureMinimum
        //ExSummary:Shows how to ensure a cell node is valid.
        Document doc = new Document();

        // Create a new table and add it to the document
        Table table = new Table(doc);
        doc.getFirstSection().getBody().appendChild(table);

        // Create a new row with a new cell and append it to the table
        Row row = new Row(doc);
        table.appendChild(row);

        Cell cell = new Cell(doc);
        row.appendChild(cell);

        // Currently, the cell does not contain any cells or nodes that can have content added to them
        Assert.assertEquals(0, cell.getChildNodes(NodeType.ANY, true).getCount());

        // Ensure the cell has at least one paragraph that we can edit
        cell.ensureMinimum();
        cell.getFirstParagraph().appendChild(new Run(doc, "Hello world!"));
        //ExEnd

        Assert.assertEquals(2, cell.getChildNodes(NodeType.ANY, true).getCount());
    }

    @Test
    public void setOutlineBorders() throws Exception {
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
        table.setBorder(BorderType.LEFT, LineStyle.SINGLE, 1.5, Color.GREEN, true);
        table.setBorder(BorderType.RIGHT, LineStyle.SINGLE, 1.5, Color.GREEN, true);
        table.setBorder(BorderType.TOP, LineStyle.SINGLE, 1.5, Color.GREEN, true);
        table.setBorder(BorderType.BOTTOM, LineStyle.SINGLE, 1.5, Color.GREEN, true);

        // Fill the cells with a light green solid color
        table.setShading(TextureIndex.TEXTURE_SOLID, Color.GREEN, Color.GREEN);

        doc.save(getArtifactsDir() + "Table.SetOutlineBorders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.SetOutlineBorders.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(TableAlignment.CENTER, table.getAlignment());

        BorderCollection borders = table.getFirstRow().getRowFormat().getBorders();

        Assert.assertEquals(Color.GREEN.getRGB(), borders.getTop().getColor().getRGB());
        Assert.assertEquals(Color.GREEN.getRGB(), borders.getLeft().getColor().getRGB());
        Assert.assertEquals(Color.GREEN.getRGB(), borders.getRight().getColor().getRGB());
        Assert.assertEquals(Color.GREEN.getRGB(), borders.getBottom().getColor().getRGB());
        Assert.assertNotEquals(Color.GREEN.getRGB(), borders.getHorizontal().getColor().getRGB());
        Assert.assertNotEquals(Color.GREEN.getRGB(), borders.getVertical().getColor().getRGB());
        Assert.assertEquals(Color.GREEN.getRGB(), table.getFirstRow().getFirstCell().getCellFormat().getShading().getForegroundPatternColor().getRGB());
    }

    @Test
    public void setTableBorders() throws Exception {
        //ExStart
        //ExFor:Table.SetBorders
        //ExSummary:Shows how to build a table with all borders enabled (grid).
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Clear any existing borders from the table
        table.clearBorders();

        // Set a green border around and inside the table
        table.setBorders(LineStyle.SINGLE, 1.5, Color.GREEN);

        doc.save(getArtifactsDir() + "Table.SetAllBorders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.SetAllBorders.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(Color.GREEN.getRGB(), table.getFirstRow().getRowFormat().getBorders().getTop().getColor().getRGB());
        Assert.assertEquals(Color.GREEN.getRGB(), table.getFirstRow().getRowFormat().getBorders().getLeft().getColor().getRGB());
        Assert.assertEquals(Color.GREEN.getRGB(), table.getFirstRow().getRowFormat().getBorders().getRight().getColor().getRGB());
        Assert.assertEquals(Color.GREEN.getRGB(), table.getFirstRow().getRowFormat().getBorders().getBottom().getColor().getRGB());
        Assert.assertEquals(Color.GREEN.getRGB(), table.getFirstRow().getRowFormat().getBorders().getHorizontal().getColor().getRGB());
        Assert.assertEquals(Color.GREEN.getRGB(), table.getFirstRow().getRowFormat().getBorders().getVertical().getColor().getRGB());
    }

    @Test
    public void rowFormat() throws Exception {
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

        doc.save(getArtifactsDir() + "Table.RowFormat.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.RowFormat.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(LineStyle.NONE, table.getFirstRow().getRowFormat().getBorders().getLineStyle());
        Assert.assertEquals(HeightRule.AUTO, table.getFirstRow().getRowFormat().getHeightRule());
        Assert.assertTrue(table.getFirstRow().getRowFormat().getAllowBreakAcrossPages());
    }

    @Test
    public void cellFormat() throws Exception {
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
        firstCell.getCellFormat().getShading().setForegroundPatternColor(Color.GREEN);

        doc.save(getArtifactsDir() + "Table.CellFormat.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.CellFormat.docx");

        table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Assert.assertEquals(30.0, table.getFirstRow().getFirstCell().getCellFormat().getWidth());
        Assert.assertEquals(TextOrientation.DOWNWARD, table.getFirstRow().getFirstCell().getCellFormat().getOrientation());
        Assert.assertEquals(Color.GREEN.getRGB(), table.getFirstRow().getFirstCell().getCellFormat().getShading().getForegroundPatternColor().getRGB());
    }

    @Test
    public void getDistance() throws Exception {
        //ExStart
        //ExFor:Table.DistanceBottom
        //ExFor:Table.DistanceLeft
        //ExFor:Table.DistanceRight
        //ExFor:Table.DistanceTop
        //ExSummary:Shows the minimum distance operations between table boundaries and text.
        Document doc = new Document(getMyDir() + "Table wrapped by text.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(table.getDistanceTop(), 25.9d);
        Assert.assertEquals(table.getDistanceBottom(), 25.9d);
        Assert.assertEquals(table.getDistanceLeft(), 17.3d);
        Assert.assertEquals(table.getDistanceRight(), 17.3d);
        //ExEnd
    }

    @Test
    public void borders() throws Exception {
        //ExStart
        //ExFor:Table.ClearBorders
        //ExSummary:Shows how to remove all borders from a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a table
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Hello world!");
        builder.endTable();

        // Set a color/thickness for the top border of the first row and verify the values
        Border topBorder = table.getFirstRow().getRowFormat().getBorders().getByBorderType(BorderType.TOP);
        table.setBorder(BorderType.TOP, LineStyle.DOUBLE, 1.5, Color.RED, true);

        Assert.assertEquals(1.5d, topBorder.getLineWidth());
        Assert.assertEquals(Color.RED.getRGB(), topBorder.getColor().getRGB());
        Assert.assertEquals(LineStyle.DOUBLE, topBorder.getLineStyle());

        // Clear the borders all cells in the table
        table.clearBorders();
        doc.save(getArtifactsDir() + "Table.ClearBorders.docx");

        // Upon re-opening the saved document, the new border attributes can be verified
        doc = new Document(getArtifactsDir() + "Table.ClearBorders.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        topBorder = table.getFirstRow().getRowFormat().getBorders().getByBorderType(BorderType.TOP);

        Assert.assertEquals(0.0d, topBorder.getLineWidth());
        Assert.assertEquals(0, topBorder.getColor().getRGB());
        Assert.assertEquals(LineStyle.NONE, topBorder.getLineStyle());
        //ExEnd
    }

    @Test
    public void replaceCellText() throws Exception {
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

        doc.save(getArtifactsDir() + "Table.ReplaceCellText.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.ReplaceCellText.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals("Eggs\u000730\u0007\u0007Potatoes\u000720", table.getText().trim());
    }

    @Test(enabled = false)
    public void printTableRange() throws Exception {
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
    public void cloneTable() throws Exception {
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
        Assert.assertEquals(doc.getChildNodes(NodeType.TABLE, true).getCount(), 3);
        Assert.assertEquals(tableClone.getRange().getText(), table.getRange().getText());

        for (Cell cell : (Iterable<Cell>) tableClone.getChildNodes(NodeType.CELL, true)) {
            cell.removeAllChildren();
        }

        Assert.assertEquals(tableClone.toString(SaveFormat.TEXT).trim(), "");
    }

    @Test
    public void disableBreakAcrossPages() throws Exception {
        //ExStart
        //ExFor:RowFormat.AllowBreakAcrossPages
        //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
        // Disable breaking across pages for all rows in the table
        Document doc = new Document(getMyDir() + "Table spanning two pages.docx");

        // Retrieve the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        for (Row row : table) {
            row.getRowFormat().setAllowBreakAcrossPages(false);
        }

        doc.save(getArtifactsDir() + "Table.DisableBreakAcrossPages.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.DisableBreakAcrossPages.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertFalse(table.getFirstRow().getRowFormat().getAllowBreakAcrossPages());
        Assert.assertFalse(table.getLastRow().getRowFormat().getAllowBreakAcrossPages());
    }

    @Test(dataProvider = "allowAutoFitOnTableDataProvider")
    public void allowAutoFitOnTable(boolean allowAutoFit) throws Exception {
        //ExStart
        //ExFor:Table.AllowAutoFit
        //ExSummary:Shows how to set a table to shrink or grow each cell to accommodate its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPoints(100.0));
        builder.write(
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.AUTO);
        builder.write(
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder.endRow();
        builder.endTable();

        table.setAllowAutoFit(allowAutoFit);

        doc.save(getArtifactsDir() + "Table.AllowAutoFitOnTable.html");
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "allowAutoFitOnTableDataProvider")
    public static Object[][] allowAutoFitOnTableDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void keepTableTogether() throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.KeepWithNext
        //ExFor:Row.IsLastRow
        //ExFor:Paragraph.IsEndOfCell
        //ExFor:Paragraph.IsInCell
        //ExFor:Cell.ParentRow
        //ExFor:Cell.Paragraphs
        //ExSummary:Shows how to set a table to stay together on the same page.
        Document doc = new Document(getMyDir() + "Table spanning two pages.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Enabling KeepWithNext for every paragraph in the table except for the last ones in the last row
        // will prevent the table from being split across pages 
        for (Cell cell : (Iterable<Cell>) table.getChildNodes(NodeType.CELL, true))
            for (Paragraph para : cell.getParagraphs()) {
                Assert.assertTrue(para.isInCell());

                if (!(cell.getParentRow().isLastRow() && para.isEndOfCell()))
                    para.getParagraphFormat().setKeepWithNext(true);
            }

        doc.save(getArtifactsDir() + "Table.KeepTableTogether.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.KeepTableTogether.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        for (Paragraph para : (Iterable<Paragraph>) table.getChildNodes(NodeType.PARAGRAPH, true))
            if (para.isEndOfCell() && ((Cell) para.getParentNode()).getParentRow().isLastRow())
                Assert.assertFalse(para.getParagraphFormat().getKeepWithNext());
            else
                Assert.assertTrue(para.getParagraphFormat().getKeepWithNext());
    }

    @Test
    public void fixDefaultTableWidthsInAw105() throws Exception {
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
    public void fixDefaultTableBordersIn105() throws Exception {
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
    public void fixDefaultTableFormattingExceptionIn105() throws Exception {
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
    public void fixRowFormattingNotAppliedIn105() throws Exception {
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
    public void getIndexOfTableElements() throws Exception {
        //ExStart
        //ExFor:NodeCollection.IndexOf(Node)
        //ExSummary:Shows how to get the indexes of nodes in the collections that contain them.
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        NodeCollection allTables = doc.getChildNodes(NodeType.TABLE, true);

        Assert.assertEquals(0, allTables.indexOf(table));

        Row row = table.getRows().get(2);

        Assert.assertEquals(2, table.indexOf(row));

        Cell cell = row.getLastCell();

        Assert.assertEquals(4, row.indexOf(cell));
        //ExEnd
    }

    @Test
    public void getPreferredWidthTypeAndValue() throws Exception {
        //ExStart
        //ExFor:PreferredWidthType
        //ExFor:PreferredWidth.Type
        //ExFor:PreferredWidth.Value
        //ExSummary:Shows how to verify the preferred width type of a table cell.
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Find the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Cell firstCell = table.getFirstRow().getFirstCell();

        Assert.assertEquals(PreferredWidthType.PERCENT, firstCell.getCellFormat().getPreferredWidth().getType());
        Assert.assertEquals(11.16, firstCell.getCellFormat().getPreferredWidth().getValue());
        //ExEnd
    }

    @Test(dataProvider = "allowCellSpacingDataProvider")
    public void allowCellSpacing(boolean allowCellSpacing) throws Exception {
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

        // Set the size of padding space between cells, and the switch that enables/negates this setting
        table.setCellSpacing(3.0);
        table.setAllowCellSpacing(allowCellSpacing);

        doc.save(getArtifactsDir() + "Table.AllowCellSpacing.html");
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "allowCellSpacingDataProvider")
    public static Object[][] allowCellSpacingDataProvider() throws Exception {
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
    //ExSummary:Shows how to build a nested table without using DocumentBuilder.
    @Test //ExSkip
    public void createNestedTable() throws Exception {
        Document doc = new Document();

        // Create the outer table with three rows and four columns
        Table outerTable = createTable(doc, 3, 4, "Outer Table");
        // Add it to the document body
        doc.getFirstSection().getBody().appendChild(outerTable);

        // Create another table with two rows and two columns
        Table innerTable = createTable(doc, 2, 2, "Inner Table");
        // Add this table to the first cell of the outer table
        outerTable.getFirstRow().getFirstCell().appendChild(innerTable);

        doc.save(getArtifactsDir() + "Table.CreateNestedTable.docx");
        testCreateNestedTable(new Document(getArtifactsDir() + "Table.CreateNestedTable.docx")); //ExSkip
    }

    /**
     * Creates a new table in the document with the given dimensions and text in each cell.
     */
    private Table createTable(final Document doc, final int rowCount, final int cellCount, final String cellText) throws Exception {
        Table table = new Table(doc);

        // Create the specified number of rows
        for (int rowId = 1; rowId <= rowCount; rowId++) {
            Row row = new Row(doc);
            table.appendChild(row);

            // Create the specified number of cells for each row
            for (int cellId = 1; cellId <= cellCount; cellId++) {
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

    private void testCreateNestedTable(Document doc) {
        Table outerTable = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Table innerTable = (Table) doc.getChild(NodeType.TABLE, 1, true);

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
    public void checkCellsMerged() throws Exception {
        Document doc = new Document(getMyDir() + "Table with merged cells.docx");

        // Retrieve the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        for (Row row : table.getRows()) {
            for (Cell cell : row.getCells()) {
                System.out.println(printCellMergeType(cell));
            }
        }

        Assert.assertEquals(printCellMergeType(table.getFirstRow().getFirstCell()), "The cell at R1, C1 is vertically merged"); //ExSkip
    }

    public String printCellMergeType(final Cell cell) {
        boolean isHorizontallyMerged = cell.getCellFormat().getHorizontalMerge() != CellMerge.NONE;
        boolean isVerticallyMerged = cell.getCellFormat().getVerticalMerge() != CellMerge.NONE;
        String cellLocation = MessageFormat.format("R{0}, C{1}", cell.getParentRow().getParentTable().indexOf(cell.getParentRow()) + 1, cell.getParentRow().indexOf(cell) + 1);

        if (isHorizontallyMerged && isVerticallyMerged) {
            return MessageFormat.format("The cell at {0} is both horizontally and vertically merged", cellLocation);
        } else if (isHorizontallyMerged) {
            return MessageFormat.format("The cell at {0} is horizontally merged.", cellLocation);
        } else if (isVerticallyMerged) {
            return MessageFormat.format("The cell at {0} is vertically merged", cellLocation);
        } else {
            return MessageFormat.format("The cell at {0} is not merged", cellLocation);
        }
    }
    //ExEnd

    @Test
    public void mergeCellRange() throws Exception {
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
        for (Cell cell : (Iterable<Cell>) table.getChildNodes(NodeType.CELL, true)) {
            if (cell.getCellFormat().getHorizontalMerge() != CellMerge.NONE || cell.getCellFormat().getHorizontalMerge() != CellMerge.NONE) {
                mergedCellsCount++;
            }
        }

        Assert.assertEquals(mergedCellsCount, 4);
        Assert.assertTrue(table.getRows().get(2).getCells().get(2).getCellFormat().getHorizontalMerge() == CellMerge.FIRST);
        Assert.assertTrue(table.getRows().get(2).getCells().get(2).getCellFormat().getVerticalMerge() == CellMerge.FIRST);
        Assert.assertTrue(table.getRows().get(3).getCells().get(3).getCellFormat().getHorizontalMerge() == CellMerge.PREVIOUS);
        Assert.assertTrue(table.getRows().get(3).getCells().get(3).getCellFormat().getVerticalMerge() == CellMerge.PREVIOUS);
    }

    /**
     * Merges the range of cells found between the two specified cells both horizontally and vertically. Can span over multiple rows.
     */
    private static void mergeCells(final Cell startCell, final Cell endCell) {
        Table parentTable = startCell.getParentRow().getParentTable();

        // Find the row and cell indices for the start and end cell
        Point startCellPos = new Point(startCell.getParentRow().indexOf(startCell),
                parentTable.indexOf(startCell.getParentRow()));
        Point endCellPos = new Point(endCell.getParentRow().indexOf(endCell),
                parentTable.indexOf(endCell.getParentRow()));
        // Create the range of cells to be merged based off these indices
        // Inverse each index if the end cell if before the start cell
        Rectangle mergeRange = new Rectangle(
                Math.min(startCellPos.x, endCellPos.x),
                Math.min(startCellPos.y, endCellPos.y),
                Math.abs(endCellPos.x - startCellPos.x) + 1,
                Math.abs(endCellPos.y - startCellPos.y) + 1);

        for (Row row : parentTable.getRows()) {
            for (Cell cell : row.getCells()) {
                Point currentPos = new Point(row.indexOf(cell), parentTable.indexOf(row));

                // Check if the current cell is inside our merge range then merge it
                if (mergeRange.contains(currentPos)) {
                    if (currentPos.x == mergeRange.x) {
                        cell.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
                    } else {
                        cell.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
                    }

                    if (currentPos.y == mergeRange.y) {
                        cell.getCellFormat().setVerticalMerge(CellMerge.FIRST);
                    } else {
                        cell.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS);
                    }
                }
            }
        }
    }

    @Test
    public void combineTables() throws Exception {
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

        doc.save(getArtifactsDir() + "Table.CombineTables.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.CombineTables.docx");

        Assert.assertEquals(1, doc.getChildNodes(NodeType.TABLE, true).getCount());
        Assert.assertEquals(9, doc.getFirstSection().getBody().getTables().get(0).getRows().getCount());
        Assert.assertEquals(42, doc.getFirstSection().getBody().getTables().get(0).getChildNodes(NodeType.CELL, true).getCount());
    }

    @Test
    public void splitTable() throws Exception {
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

        do {
            currentRow = firstTable.getLastRow();
            table.prependChild(currentRow);
        } while (currentRow != row);

        doc.save(getArtifactsDir() + "Table.SplitTable.docx");

        doc = new Document(getArtifactsDir() + "Table.SplitTable.docx");
        // Test we are adding the rows in the correct order and the
        // selected row was also moved
        Assert.assertEquals(table.getFirstRow(), row);

        Assert.assertEquals(firstTable.getRows().getCount(), 2);
        Assert.assertEquals(table.getRows().getCount(), 3);
        Assert.assertEquals(doc.getChildNodes(NodeType.TABLE, true).getCount(), 3);
    }

    @Test
    public void wrapText() throws Exception {
        //ExStart
        //ExFor:Table.TextWrapping
        //ExFor:TextWrapping
        //ExSummary:Shows how to work with table text wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table and a paragraph of text after it
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.endTable();
        table.setPreferredWidth(PreferredWidth.fromPoints(300.0));

        builder.getFont().setSize(16.0);
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Set the table to wrap text around it and push it down into the paragraph below be setting the position
        table.setTextWrapping(TextWrapping.AROUND);
        table.setAbsoluteHorizontalDistance(100.0);
        table.setAbsoluteVerticalDistance(20.0);

        doc.save(getArtifactsDir() + "Table.WrapText.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.WrapText.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(TextWrapping.AROUND, table.getTextWrapping());
        Assert.assertEquals(100.0d, table.getAbsoluteHorizontalDistance());
        Assert.assertEquals(20.0d, table.getAbsoluteVerticalDistance());
    }

    @Test
    public void getFloatingTableProperties() throws Exception {
        //ExStart
        //ExFor:Table.HorizontalAnchor
        //ExFor:Table.VerticalAnchor
        //ExFor:Table.AllowOverlap
        //ExFor:ShapeBase.AllowOverlap
        //ExSummary:Shows how to work with floating tables properties.
        Document doc = new Document(getMyDir() + "Table wrapped by text.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        if (table.getTextWrapping() == TextWrapping.AROUND) {
            Assert.assertEquals(TextWrapping.AROUND, table.getTextWrapping());
            Assert.assertEquals(RelativeHorizontalPosition.MARGIN, table.getHorizontalAnchor());
            Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, table.getVerticalAnchor());
            Assert.assertEquals(false, table.getAllowOverlap());

            // Only Margin, Page, Column available in RelativeHorizontalPosition for HorizontalAnchor setter
            // The ArgumentException will be thrown for any other values
            table.setHorizontalAnchor(RelativeHorizontalPosition.COLUMN);
            // Only Margin, Page, Paragraph available in RelativeVerticalPosition for VerticalAnchor setter
            // The ArgumentException will be thrown for any other values
            table.setVerticalAnchor(RelativeVerticalPosition.PAGE);

            Assert.assertEquals(RelativeHorizontalPosition.COLUMN, table.getHorizontalAnchor()); //ExSkip
            Assert.assertEquals(RelativeVerticalPosition.PAGE, table.getVerticalAnchor()); //ExSkip
        }
        //ExEnd
    }

    @Test
    public void changeFloatingTableProperties() throws Exception {
        //ExStart
        //ExFor:Table.RelativeHorizontalAlignment
        //ExFor:Table.RelativeVerticalAlignment
        //ExFor:Table.AbsoluteHorizontalDistance
        //ExFor:Table.AbsoluteVerticalDistance
        //ExSummary:Shows how set the location of floating tables.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Table 1, cell 1");
        builder.endTable();
        table.setPreferredWidth(PreferredWidth.fromPoints(300.0));

        // We can set the table's location to a place on the page, such as the bottom right corner
        table.setRelativeVerticalAlignment(VerticalAlignment.BOTTOM);
        table.setRelativeHorizontalAlignment(HorizontalAlignment.RIGHT);

        table = builder.startTable();
        builder.insertCell();
        builder.write("Table 2, cell 1");
        builder.endTable();
        table.setPreferredWidth(PreferredWidth.fromPoints(300.0));

        // We can also set a horizontal and vertical offset from the location in the paragraph where the table was inserted 
        table.setAbsoluteVerticalDistance(50.0);
        table.setAbsoluteHorizontalDistance(100.0);

        doc.save(getArtifactsDir() + "Table.ChangeFloatingTableProperties.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.ChangeFloatingTableProperties.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(VerticalAlignment.BOTTOM, table.getRelativeVerticalAlignment());
        Assert.assertEquals(HorizontalAlignment.RIGHT, table.getRelativeHorizontalAlignment());

        table = (Table) doc.getChild(NodeType.TABLE, 1, true);

        Assert.assertEquals(50.0d, table.getAbsoluteVerticalDistance());
        Assert.assertEquals(100.0d, table.getAbsoluteHorizontalDistance());
    }

    @Test
    public void tableStyleCreation() throws Exception {
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

        TableStyle tableStyle = (TableStyle) doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.setAllowBreakAcrossPages(true);
        tableStyle.setBidi(true);
        tableStyle.setCellSpacing(5.0);
        tableStyle.setBottomPadding(20.0);
        tableStyle.setLeftPadding(5.0);
        tableStyle.setRightPadding(10.0);
        tableStyle.setTopPadding(20.0);
        tableStyle.getShading().setBackgroundPatternColor(Color.WHITE);
        tableStyle.getBorders().setColor(Color.BLACK);
        tableStyle.getBorders().setLineStyle(LineStyle.DOT_DASH);

        table.setStyle(tableStyle);

        // Some Table attributes are linked to style variables
        Assert.assertEquals(table.getBidi(), true);
        Assert.assertEquals(table.getCellSpacing(), 5.0);
        Assert.assertEquals(table.getStyleName(), "MyTableStyle1");

        doc.save(getArtifactsDir() + "Table.TableStyleCreation.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.TableStyleCreation.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertTrue(table.getBidi());
        Assert.assertEquals(5.0d, table.getCellSpacing());
        Assert.assertEquals("MyTableStyle1", table.getStyleName());
        Assert.assertEquals(0.0d, table.getBottomPadding());
        Assert.assertEquals(0.0d, table.getLeftPadding());
        Assert.assertEquals(0.0d, table.getRightPadding());
        Assert.assertEquals(0.0d, table.getTopPadding());

        tableStyle = (TableStyle) doc.getStyles().get("MyTableStyle1");

        Assert.assertTrue(tableStyle.getAllowBreakAcrossPages());
        Assert.assertTrue(tableStyle.getBidi());
        Assert.assertEquals(5.0d, tableStyle.getCellSpacing());
        Assert.assertEquals(20.0d, tableStyle.getBottomPadding());
        Assert.assertEquals(5.0d, tableStyle.getLeftPadding());
        Assert.assertEquals(10.0d, tableStyle.getRightPadding());
        Assert.assertEquals(20.0d, tableStyle.getTopPadding());
        Assert.assertEquals(Color.WHITE.getRGB(), tableStyle.getShading().getBackgroundPatternColor().getRGB());
        Assert.assertEquals(Color.BLACK.getRGB(), tableStyle.getBorders().getColor().getRGB());
        Assert.assertEquals(LineStyle.DOT_DASH, tableStyle.getBorders().getLineStyle());
    }

    @Test
    public void setTableAlignment() throws Exception {
        //ExStart
        //ExFor:TableStyle.Alignment
        //ExFor:TableStyle.LeftIndent
        //ExSummary:Shows how to set table position.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // There are two ways of horizontally aligning a table using a custom table style
        // One way is to align it to a location on the page, such as the center
        TableStyle tableStyle = (TableStyle) doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.setAlignment(TableAlignment.CENTER);
        tableStyle.getBorders().setColor(Color.BLUE);
        tableStyle.getBorders().setLineStyle(LineStyle.SINGLE);

        // Insert a table and apply the style we created to it
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Aligned to the center of the page");
        builder.endTable();
        table.setPreferredWidth(PreferredWidth.fromPoints(300.0));

        table.setStyle(tableStyle);

        // We can also set a specific left indent to the style, and apply it to the table
        tableStyle = (TableStyle) doc.getStyles().add(StyleType.TABLE, "MyTableStyle2");
        tableStyle.setLeftIndent(55.0);
        tableStyle.getBorders().setColor(Color.GREEN);
        tableStyle.getBorders().setLineStyle(LineStyle.SINGLE);

        table = builder.startTable();
        builder.insertCell();
        builder.write("Aligned according to left indent");
        builder.endTable();
        table.setPreferredWidth(PreferredWidth.fromPoints(300.0));

        table.setStyle(tableStyle);

        doc.save(getArtifactsDir() + "Table.TableStyleCreation.docx");
        //ExEnd
    }

    @Test
    public void conditionalStyles() throws Exception {
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

        // Create a table
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

        // Create a custom table style
        TableStyle tableStyle = (TableStyle) doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");

        // Conditional styles are formatting changes that affect only some of the cells of the table based on a predicate,
        // such as the cells being in the last row
        // We can access these conditional styles by style type like this
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.FIRST_ROW).getShading().setBackgroundPatternColor(Color.BLUE);

        // The same conditional style can be accessed by index
        tableStyle.getConditionalStyles().get(0).getBorders().setColor(Color.BLACK);
        tableStyle.getConditionalStyles().get(0).getBorders().setLineStyle(LineStyle.DOT_DASH);
        Assert.assertEquals(ConditionalStyleType.FIRST_ROW, tableStyle.getConditionalStyles().get(0).getType());

        // It can also be found in the ConditionalStyles collection as an attribute
        tableStyle.getConditionalStyles().getFirstRow().getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        // Apply padding and text formatting to conditional styles 
        tableStyle.getConditionalStyles().getLastRow().setBottomPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setLeftPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setRightPadding(10.0);
        tableStyle.getConditionalStyles().getLastRow().setTopPadding(10.0);
        tableStyle.getConditionalStyles().getLastColumn().getFont().setBold(true);

        // List all possible style conditions
        Iterator<ConditionalStyle> enumerator = tableStyle.getConditionalStyles().iterator();
        while (enumerator.hasNext()) {
            ConditionalStyle currentStyle = enumerator.next();
            if (currentStyle != null) System.out.println(currentStyle.getType());
        }


        // Apply conditional style to the table
        table.setStyle(tableStyle);

        // Changes to the first row are enabled by the table's style options be default,
        // but need to be manually enabled for some other parts, such as the last column/row
        table.setStyleOptions(table.getStyleOptions() | TableStyleOptions.LAST_ROW | TableStyleOptions.LAST_COLUMN);

        doc.save(getArtifactsDir() + "Table.ConditionalStyles.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.ConditionalStyles.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(TableStyleOptions.DEFAULT | TableStyleOptions.LAST_ROW | TableStyleOptions.LAST_COLUMN, table.getStyleOptions());
        ConditionalStyleCollection conditionalStyles = ((TableStyle) doc.getStyles().get("MyTableStyle1")).getConditionalStyles();

        Assert.assertEquals(ConditionalStyleType.FIRST_ROW, conditionalStyles.get(0).getType());
        Assert.assertEquals(Color.BLUE.getRGB(), conditionalStyles.get(0).getShading().getBackgroundPatternColor().getRGB());
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
    public void clearTableStyleFormatting() throws Exception {
        //ExStart
        //ExFor:ConditionalStyle.ClearFormatting
        //ExFor:ConditionalStyleCollection.ClearFormatting
        //ExSummary:Shows how to reset conditional table styles.
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

        TableStyle tableStyle = (TableStyle) doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.getConditionalStyles().getFirstRow().getBorders().setColor(Color.RED);
        tableStyle.getConditionalStyles().getLastRow().getBorders().setColor(Color.BLUE);

        // Conditional styles can be cleared for specific parts of the table 
        tableStyle.getConditionalStyles().get(0).clearFormatting();
        Assert.assertEquals(tableStyle.getConditionalStyles().getFirstRow().getBorders().getColor().getRGB(), 0);

        // Also, they can be cleared for the entire table
        tableStyle.getConditionalStyles().clearFormatting();
        Assert.assertEquals(tableStyle.getConditionalStyles().getLastRow().getBorders().getColor().getRGB(), 0);
        //ExEnd
    }

    @Test
    public void alternatingRowStyles() throws Exception {
        //ExStart
        //ExFor:TableStyle.ColumnStripe
        //ExFor:TableStyle.RowStripe
        //ExSummary:Shows how to create conditional table styles that alternate between rows.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The conditional style of a table can be configured to apply a different color to the row/column,
        // based on whether the row/column is even or odd, creating an alternating color pattern
        // We can also apply a number n to the row/column banding, meaning that the color alternates after every n rows/columns instead of one 
        // Create a table where the columns will be banded by single columns and rows will banded in threes
        Table table = builder.startTable();

        for (int i = 0; i < 15; i++) {
            for (int j = 0; j < 4; j++) {
                builder.insertCell();
                builder.writeln(MessageFormat.format("{0} column.", (j % 2 == 0 ? "Even" : "Odd")));
                builder.write(MessageFormat.format("Row banding {0}.", (i % 3 == 0 ? "start" : "continuation")));
            }
            builder.endRow();
        }

        builder.endTable();

        // Set a line style for all the borders of the table
        TableStyle tableStyle = (TableStyle) doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.getBorders().setColor(Color.BLACK);
        tableStyle.getBorders().setLineStyle(LineStyle.DOUBLE);

        // Set the two colors which will alternate over every 3 rows
        tableStyle.setRowStripe(3);
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.ODD_ROW_BANDING).getShading().setBackgroundPatternColor(Color.BLUE);
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.EVEN_ROW_BANDING).getShading().setBackgroundPatternColor(Color.CYAN);

        // Set a color to apply to every even column, which will override any custom row coloring
        tableStyle.setColumnStripe(1);
        tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.EVEN_COLUMN_BANDING).getShading().setBackgroundPatternColor(Color.RED);

        // Apply the style to the table
        table.setStyle(tableStyle);

        // Row bands are automatically enabled, but column banding needs to be enabled manually like this
        // Row coloring will only be overridden if the column banding has been explicitly set a color
        table.setStyleOptions(table.getStyleOptions() | TableStyleOptions.COLUMN_BANDS);

        doc.save(getArtifactsDir() + "Table.AlternatingRowStyles.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Table.AlternatingRowStyles.docx");
        table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        tableStyle = (TableStyle) doc.getStyles().get("MyTableStyle1");

        Assert.assertEquals(tableStyle, table.getStyle());
        Assert.assertEquals(table.getStyleOptions() | TableStyleOptions.COLUMN_BANDS, table.getStyleOptions());

        Assert.assertEquals(Color.BLACK.getRGB(), tableStyle.getBorders().getColor().getRGB());
        Assert.assertEquals(LineStyle.DOUBLE, tableStyle.getBorders().getLineStyle());
        Assert.assertEquals(3, tableStyle.getRowStripe());
        Assert.assertEquals(Color.BLUE.getRGB(), tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.ODD_ROW_BANDING).getShading().getBackgroundPatternColor().getRGB());
        Assert.assertEquals(Color.CYAN.getRGB(), tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.EVEN_ROW_BANDING).getShading().getBackgroundPatternColor().getRGB());
        Assert.assertEquals(1, tableStyle.getColumnStripe());
        Assert.assertEquals(Color.RED.getRGB(), tableStyle.getConditionalStyles().getByConditionalStyleType(ConditionalStyleType.EVEN_COLUMN_BANDING).getShading().getBackgroundPatternColor().getRGB());
    }

    @Test
    public void convertToHorizontallyMergedCells() throws Exception {
        //ExStart
        //ExFor:Table.ConvertToHorizontallyMergedCells
        //ExSummary:Shows how to convert cells horizontally merged by width to cells merged by CellFormat.HorizontalMerge.
        Document doc = new Document(getMyDir() + "Table with merged cells.docx");

        // MS Word does not write merge flags anymore, they define merged cells by its width
        // So AW by default define only 5 cells in a row and all of it didn't have horizontal merge flag
        Table table = doc.getFirstSection().getBody().getTables().get(0);
        Row row = table.getRows().get(0);
        Assert.assertEquals(row.getCells().getCount(), 5);

        // To resolve this inconvenience, we have added new public method to convert cells which are horizontally merged
        // by its width to the cell horizontally merged by flags. Thus now we have 7 cells and some of them have
        // horizontal merge value
        table.convertToHorizontallyMergedCells();
        row = table.getRows().get(0);
        Assert.assertEquals(row.getCells().getCount(), 7);

        Assert.assertEquals(row.getCells().get(0).getCellFormat().getHorizontalMerge(), CellMerge.NONE);
        Assert.assertEquals(row.getCells().get(1).getCellFormat().getHorizontalMerge(), CellMerge.FIRST);
        Assert.assertEquals(row.getCells().get(2).getCellFormat().getHorizontalMerge(), CellMerge.PREVIOUS);
        Assert.assertEquals(row.getCells().get(3).getCellFormat().getHorizontalMerge(), CellMerge.NONE);
        Assert.assertEquals(row.getCells().get(4).getCellFormat().getHorizontalMerge(), CellMerge.FIRST);
        Assert.assertEquals(row.getCells().get(5).getCellFormat().getHorizontalMerge(), CellMerge.PREVIOUS);
        Assert.assertEquals(row.getCells().get(6).getCellFormat().getHorizontalMerge(), CellMerge.NONE);
        //ExEnd
    }
}
