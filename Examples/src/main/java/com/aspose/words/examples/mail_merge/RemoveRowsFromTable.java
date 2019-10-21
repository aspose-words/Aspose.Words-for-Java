package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.MailMergeCleanupOptions;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataSet;

public class RemoveRowsFromTable {

    private static final String dataDir = Utils.getDataDir(RemoveRowsFromTable.class);

    public static void main(String[] args) throws Exception {
        //Exstart:RemoveRowsFromTable
        Document doc = new Document(dataDir + "RemoveTableRows.doc");
        DataSet data = new DataSet();
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS | MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS | MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS);
        doc.getMailMerge().setMergeDuplicateRegions(true);
        doc.getMailMerge().executeWithRegions(data);
        doc.save(dataDir + "RemoveTableRows_Out.doc");
        //ExEnd:RemoveRowsFromTable
    }


}