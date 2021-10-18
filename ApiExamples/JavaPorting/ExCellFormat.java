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
import com.aspose.words.DocumentBuilder;
import com.aspose.words.CellMerge;
import com.aspose.words.Table;
import org.testng.Assert;
import com.aspose.ms.System.msString;
import com.aspose.words.Cell;


@Test
public class ExCellFormat extends ApiExampleBase
{
    @Test
    public void verticalMerge() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.EndRow
        //ExFor:CellMerge
        //ExFor:CellFormat.VerticalMerge
        //ExSummary:Shows how to merge table cells vertically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a cell into the first column of the first row.
        // This cell will be the first in a range of vertically merged cells.
        builder.insertCell();
        builder.getCellFormat().setVerticalMerge(CellMerge.FIRST);
        builder.write("Text in merged cells.");

        // Insert a cell into the second column of the first row, then end the row.
        // Also, configure the builder to disable vertical merging in created cells.
        builder.insertCell();
        builder.getCellFormat().setVerticalMerge(CellMerge.NONE);
        builder.write("Text in unmerged cell.");
        builder.endRow();

        // Insert a cell into the first column of the second row. 
        // Instead of adding text contents, we will merge this cell with the first cell that we added directly above.
        builder.insertCell();
        builder.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS);

        // Insert another independent cell in the second column of the second row.
        builder.insertCell();
        builder.getCellFormat().setVerticalMerge(CellMerge.NONE);
        builder.write("Text in unmerged cell.");
        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "CellFormat.VerticalMerge.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "CellFormat.VerticalMerge.docx");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(CellMerge.FIRST, table.getRows().get(0).getCells().get(0).getCellFormat().getVerticalMerge());
        Assert.assertEquals(CellMerge.PREVIOUS, table.getRows().get(1).getCells().get(0).getCellFormat().getVerticalMerge());
        Assert.assertEquals("Text in merged cells.", msString.trim(table.getRows().get(0).getCells().get(0).getText(), '\u0007'));
        Assert.assertNotEquals(table.getRows().get(0).getCells().get(0).getText(), table.getRows().get(1).getCells().get(0).getText());
    }

    @Test
    public void horizontalMerge() throws Exception
    {
        //ExStart
        //ExFor:CellMerge
        //ExFor:CellFormat.HorizontalMerge
        //ExSummary:Shows how to merge table cells horizontally.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a cell into the first column of the first row.
        // This cell will be the first in a range of horizontally merged cells.
        builder.insertCell();
        builder.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
        builder.write("Text in merged cells.");

        // Insert a cell into the second column of the first row. Instead of adding text contents,
        // we will merge this cell with the first cell that we added directly to the left.
        builder.insertCell();
        builder.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
        builder.endRow();

        // Insert two more unmerged cells to the second row.
        builder.getCellFormat().setHorizontalMerge(CellMerge.NONE);
        builder.insertCell();
        builder.write("Text in unmerged cell.");
        builder.insertCell();
        builder.write("Text in unmerged cell.");
        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "CellFormat.HorizontalMerge.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "CellFormat.HorizontalMerge.docx");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(1, table.getRows().get(0).getCells().getCount());
        Assert.assertEquals(CellMerge.NONE, table.getRows().get(0).getCells().get(0).getCellFormat().getHorizontalMerge());
        Assert.assertEquals("Text in merged cells.", msString.trim(table.getRows().get(0).getCells().get(0).getText(), '\u0007'));
    }

    @Test
    public void padding() throws Exception
    {
        //ExStart
        //ExFor:CellFormat.SetPaddings
        //ExSummary:Shows how to pad the contents of a cell with whitespace.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a padding distance (in points) between the border and the text contents
        // of each table cell we create with the document builder. 
        builder.getCellFormat().setPaddings(5.0, 10.0, 40.0, 50.0);

        // Create a table with one cell whose contents will have whitespace padding.
        builder.startTable();
        builder.insertCell();
        builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                      "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        doc.save(getArtifactsDir() + "CellFormat.Padding.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "CellFormat.Padding.docx");

        Table table = doc.getFirstSection().getBody().getTables().get(0);
        Cell cell = table.getRows().get(0).getCells().get(0);

        Assert.assertEquals(5, cell.getCellFormat().getLeftPadding());
        Assert.assertEquals(10, cell.getCellFormat().getTopPadding());
        Assert.assertEquals(40, cell.getCellFormat().getRightPadding());
        Assert.assertEquals(50, cell.getCellFormat().getBottomPadding());
    }
}
