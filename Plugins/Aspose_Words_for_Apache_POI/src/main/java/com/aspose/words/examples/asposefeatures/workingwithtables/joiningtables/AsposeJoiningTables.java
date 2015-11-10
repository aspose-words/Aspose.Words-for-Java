package com.aspose.words.examples.asposefeatures.workingwithtables.joiningtables;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class AsposeJoiningTables
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeJoiningTables.class);

        // Load the document.
        Document doc = new Document(dataDir + "tableDoc.doc");

        // Get the first and second table in the document.
        // The rows from the second table will be appended to the end of the first table.
        Table firstTable = (Table)doc.getChild(NodeType.TABLE, 0, true);
        Table secondTable = (Table)doc.getChild(NodeType.TABLE, 1, true);

        // Append all rows from the current table to the next.
        // Due to the design of tables even tables with different cell count and widths can be joined into one table.
        while (secondTable.hasChildNodes())
            firstTable.getRows().add(secondTable.getFirstRow());

        // Remove the empty table container.
        secondTable.remove();

        doc.save(dataDir + "AsposeJoinTables.doc");

        System.out.println("Done.");
    }
}