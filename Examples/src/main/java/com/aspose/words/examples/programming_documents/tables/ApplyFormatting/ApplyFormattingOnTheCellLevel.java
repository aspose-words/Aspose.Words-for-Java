package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;

public class ApplyFormattingOnTheCellLevel {

    private static final String dataDir = Utils.getSharedDataDir(ApplyFormattingOnTheCellLevel.class) + "Tables/";

    public static void main(String[] args) throws Exception {
        //ExStart:ApplyFormattingOnTheCellLevel
        Document doc = new Document(dataDir + "Table.Document.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Retrieve the first cell in the table.
        Cell firstCell = table.getFirstRow().getFirstCell();

        // Modify some row level properties.
        firstCell.getCellFormat().setWidth(30); // in points
        firstCell.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        firstCell.getCellFormat().getShading().setForegroundPatternColor(Color.GREEN);
        //ExEnd:ApplyFormattingOnTheCellLevel
    }
}
