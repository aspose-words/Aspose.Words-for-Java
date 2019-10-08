package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class CheckCellsMerged {

    private static final String dataDir = Utils.getSharedDataDir(CheckCellsMerged.class) + "Tables/";

    public static void main(String[] args) throws Exception {
        //ExStart:CheckCellsMerged
        Document doc = new Document(dataDir + "Table.MergedCells.doc");

        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        for (Row row : table.getRows()) {
            for (Cell cell : row.getCells()) {
                System.out.println(printCellMergeType(cell));
            }
        }
        //ExEnd:CheckCellsMerged
    }

    //ExStart:printCellMergeType
    public static String printCellMergeType(Cell cell) {
        boolean isHorizontallyMerged = cell.getCellFormat().getHorizontalMerge() != CellMerge.NONE;
        boolean isVerticallyMerged = cell.getCellFormat().getVerticalMerge() != CellMerge.NONE;
        String cellLocation = "R" + (cell.getParentRow().getParentTable().indexOf(cell.getParentRow()) + 1) +
                ", C" + (cell.getParentRow().indexOf(cell) + 1);

        if (isHorizontallyMerged && isVerticallyMerged)
            return "The cell at " + cellLocation + " is both horizontally and vertically merged";
        else if (isHorizontallyMerged)
            return "The cell at " + cellLocation + " is horizontally merged.";
        else if (isVerticallyMerged)
            return "The cell at " + cellLocation + " is vertically merged";
        else
            return "The cell at " + cellLocation + " is not merged";
    }//ExEnd:printCellMergeType
}
