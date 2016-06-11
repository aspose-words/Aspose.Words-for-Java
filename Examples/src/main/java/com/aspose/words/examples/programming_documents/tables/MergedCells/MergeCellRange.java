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


public class MergeCellRange {
    public static void main(String[] args) throws Exception {
        //TODO:MergeCells
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(MergeCellRange.class);

        Document doc = new Document();
        // Retrieve the first table in the body of the first section.
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        // We want to merge the range of cells found inbetween these two cells.
        Cell cellStartRange = table.getRows().get(2).getCells().get(2);
        Cell cellEndRange = table.getRows().get(3).getCells().get(3);

        // Merge all the cells between the two specified cells into one.


        //MergeCells(cellStartRange, cellEndRange);
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}