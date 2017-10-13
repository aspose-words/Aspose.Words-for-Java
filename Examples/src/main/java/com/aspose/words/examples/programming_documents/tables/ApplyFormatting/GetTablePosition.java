package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.TextWrapping;
import com.aspose.words.examples.Utils;


/**
 * Created by Home on 10/13/2017.
 */
public class GetTablePosition {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(GetTablePosition.class);
        // ExStart:GetTablePosition
        Document doc = new Document(dataDir + "Table.Document.doc");

        // Retrieve the first table in the document.
        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

        if (table.getTextWrapping() == TextWrapping.AROUND)
        {
            System.out.println("RelativeHorizontalAlignment: "+table.getRelativeHorizontalAlignment());
            System.out.println("RelativeVerticalAlignment: "+table.getRelativeVerticalAlignment());
        }
        else
        {
            System.out.println("Alignment: "+table.getAlignment());
        }

        // ExEnd:GetTablePosition
        System.out.println("\nTable position fetched successfully.");
    }
}
