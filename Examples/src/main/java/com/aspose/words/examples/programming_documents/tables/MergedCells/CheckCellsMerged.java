/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.MergedCells;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class CheckCellsMerged {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(CheckCellsMerged.class);

        Document doc = new Document(dataDir + "Table.TableAcrossPage.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        for (Row row : table.getRows()) {
            for (Cell cell : row.getCells()) {
                boolean isHorizontalMerged = (cell.getCellFormat().getHorizontalMerge()) != CellMerge.NONE;
                boolean isVerticalMerged = cell.getCellFormat().getVerticalMerge() != CellMerge.NONE;
                String cellLocation = String.format("R%s, C%s", cell.getParentRow().getParentTable().indexOf(cell.getParentRow()) + 1, cell.getParentRow().indexOf(cell) + 1);

                if (isHorizontalMerged && isVerticalMerged)
                    System.out.println(String.format("The cell at %s is both horizontally and vertically merged", cellLocation));
                else if (isHorizontalMerged)
                    System.out.println(String.format("The cell at %s is horizontally merged.", cellLocation));
                else if (isVerticalMerged)
                    System.out.println(String.format("The cell at %s is vertically merged", cellLocation));
                else
                    System.out.println(String.format("The cell at %s is not merged", cellLocation));
            }
        }
        //ExEnd:1
    }
}
