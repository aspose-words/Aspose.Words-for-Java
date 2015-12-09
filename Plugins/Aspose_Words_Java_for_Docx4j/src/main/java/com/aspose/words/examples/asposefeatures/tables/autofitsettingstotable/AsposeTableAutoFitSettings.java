package com.aspose.words.examples.asposefeatures.tables.autofitsettingstotable;

import com.aspose.words.AutoFitBehavior;
import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class AsposeTableAutoFitSettings
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeTableAutoFitSettings.class);

        // Open the document
        Document doc = new Document(dataDir + "tableDoc.doc");

        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);
        // Autofit the first table to the page width.
        table.autoFit(AutoFitBehavior.AUTO_FIT_TO_WINDOW);

        Table table2 = (Table)doc.getChild(NodeType.TABLE, 1, true);
        // Auto fit the table to the cell contents
        table2.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

        // Save the document to disk.
        doc.save(dataDir + "AsposeAutoFitTable_Out.doc");

        System.out.println("Process Completed Successfully");
    }
}