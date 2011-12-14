//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.CellMerge;


public class ExCellFormat extends ExBase
{
    @Test
    public void verticalMerge() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertCell
        //ExFor:DocumentBuilder.EndRow
        //ExFor:CellMerge
        //ExFor:CellFormat.VerticalMerge
        //ExId:VerticalMerge
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
        //ExEnd
    }

    @Test
    public void horizontalMerge() throws Exception
    {
        //ExStart
        //ExFor:CellMerge
        //ExFor:CellFormat.HorizontalMerge
        //ExId:HorizontalMerge
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
        //ExEnd
    }
}

