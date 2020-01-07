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

import java.io.ByteArrayOutputStream;

public class ExCellFormat extends ApiExampleBase {
    @Test
    public void verticalMerge() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertCell
        //ExFor:DocumentBuilder.EndRow
        //ExFor:CellMerge
        //ExFor:CellFormat.VerticalMerge
        //ExSummary:Creates a table with two columns with cells merged vertically in the first column.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertCell();
        builder.getCellFormat().setVerticalMerge(CellMerge.FIRST);
        builder.write("Text in merged cells.");

        builder.insertCell();
        builder.getCellFormat().setVerticalMerge(CellMerge.NONE);
        builder.write("Text in one cell");
        builder.endRow();

        builder.insertCell();
        // This cell is vertically merged to the cell above and should be empty.
        builder.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS);

        builder.insertCell();
        builder.getCellFormat().setVerticalMerge(CellMerge.NONE);
        builder.write("Text in another cell");
        builder.endRow();
        builder.endTable();
        //ExEnd
    }

    @Test
    public void horizontalMerge() throws Exception {
        //ExStart
        //ExFor:CellMerge
        //ExFor:CellFormat.HorizontalMerge
        //ExSummary:Creates a table with two rows with cells in the first row horizontally merged.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertCell();
        builder.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
        builder.write("Text in merged cells.");

        builder.insertCell();
        // This cell is merged to the previous and should be empty.
        builder.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
        builder.endRow();

        builder.insertCell();
        builder.getCellFormat().setHorizontalMerge(CellMerge.NONE);
        builder.write("Text in one cell.");

        builder.insertCell();
        builder.write("Text in another cell.");
        builder.endRow();
        builder.endTable();
        //ExEnd
    }

    @Test
    public void setCellPaddings() throws Exception {
        //ExStart
        //ExFor:CellFormat.SetPaddings
        //ExSummary:Shows how to set paddings to a table cell.
        DocumentBuilder builder = new DocumentBuilder();

        builder.startTable();
        builder.getCellFormat().setWidth(300.0);
        builder.getCellFormat().setPaddings(5.0, 10.0, 40.0, 50.0);

        builder.getRowFormat().setHeightRule(HeightRule.EXACTLY);
        builder.getRowFormat().setHeight(50.0);

        builder.insertCell();
        builder.write("Row 1, Col 1");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        builder.getDocument().save(dstStream, SaveFormat.DOCX);

        Table table = (Table) builder.getDocument().getChild(NodeType.TABLE, 0, true);

        Cell cell = table.getRows().get(0).getCells().get(0);

        Assert.assertEquals(cell.getCellFormat().getLeftPadding(), 5.0);
        Assert.assertEquals(cell.getCellFormat().getTopPadding(), 10.0);
        Assert.assertEquals(cell.getCellFormat().getRightPadding(), 40.0);
        Assert.assertEquals(cell.getCellFormat().getBottomPadding(), 50.0);
        //ExEnd
    }
}
