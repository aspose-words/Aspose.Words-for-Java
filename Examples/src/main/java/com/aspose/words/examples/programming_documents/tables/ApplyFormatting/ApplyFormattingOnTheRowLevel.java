package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class ApplyFormattingOnTheRowLevel {

    private static final String dataDir = Utils.getSharedDataDir(ApplyFormattingOnTheRowLevel.class) + "Tables/";

    public static void main(String[] args) throws Exception {

        //ExStart:ApplyFormattingOnTheRowLevel
        Document doc = new Document(dataDir + "Table.Document.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Retrieve the first row in the table.
        Row firstRow = table.getFirstRow();

        // Modify some row level properties.
        firstRow.getRowFormat().getBorders().setLineStyle(LineStyle.NONE);
        firstRow.getRowFormat().setHeightRule(HeightRule.AUTO);
        firstRow.getRowFormat().setAllowBreakAcrossPages(true);
        //ExEnd:ApplyFormattingOnTheRowLevel
    }
}
