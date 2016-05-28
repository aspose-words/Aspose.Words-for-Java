/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.AddRemoveColumn;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;


public class InsertBlankColumn {
    public static void main(String[] args) throws Exception {

        //TODO
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertBlankColumn.class);
        Document doc = new Document(dataDir + "Table.SimpleTable.doc");
        // Get the first table in the document.
        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);
        // Get the second column in the table.
       // Column column = Column.FromIndex(table, 0);
        // Print the plain text of the column to the screen.
      //  Console.WriteLine(column.ToTxt());
        // Create a new column to the left of this column.
        // This is the same as using the "Insert Column Before" command in Microsoft Word.
//        Column newColumn = column.InsertColumnBefore();

        // Add some text to each of the column cells.
       // foreach (Cell cell in newColumn.Cells)
       // cell.FirstParagraph.AppendChild(new Run(doc, "Column Text " + newColumn.IndexOf(cell)));

    //    doc.save(dataDir + "output.doc");

        //ExEnd:1
    }
}