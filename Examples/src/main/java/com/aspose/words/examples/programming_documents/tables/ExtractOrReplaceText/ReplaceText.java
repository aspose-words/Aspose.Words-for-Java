package com.aspose.words.examples.programming_documents.tables.ExtractOrReplaceText;

import com.aspose.words.Document;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class ReplaceText {

    private static final String dataDir = Utils.getSharedDataDir(ReplaceText.class) + "Tables/";

    public static void main(String[] args) throws Exception {
        //ExStart:ReplaceText
        Document doc = new Document(dataDir + "Table.SimpleTable.doc");

        // Get the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        FindReplaceOptions opts = new FindReplaceOptions();
        opts.setMatchCase(true);
        opts.setFindWholeWordsOnly(true);

        // Replace any instances of our string in the entire table.
        table.getRange().replace("Carrots", "Eggs", opts);

        // Replace any instances of our string in the last cell of the table only.
        table.getLastRow().getLastCell().getRange().replace("50", "20", opts);

        doc.save(dataDir + "Table.ReplaceCellText Out.docx");
        //ExEnd:ReplaceText
    }
}