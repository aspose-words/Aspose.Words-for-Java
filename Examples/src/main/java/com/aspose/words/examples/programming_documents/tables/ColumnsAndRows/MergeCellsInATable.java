package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import com.aspose.words.CellMerge;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;

public class MergeCellsInATable {

    public static void main(String[] args) {

    }

    //ExStart:mergeCellsHorizontally
    public static void mergeCellsHorizontally() throws Exception {

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
    }
    //ExEnd:mergeCellsHorizontally

    //ExStart:mergeCellsVertically
    public static void mergeCellsVertically() throws Exception {
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
    }
    //ExEnd:mergeCellsVertically
}
