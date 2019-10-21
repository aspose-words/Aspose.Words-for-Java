package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import com.aspose.words.Document;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class ConvertToHorizontallyMergedCells {

    private static final String dataDir = Utils.getSharedDataDir(ConvertToHorizontallyMergedCells.class) + "Tables/";

    public static void main(String[] args) throws Exception {
        // ExStart:ConvertToHorizontallyMergedCells
        Document doc = new Document();

        Table table = doc.getFirstSection().getBody().getTables().get(0);
        table.convertToHorizontallyMergedCells();   // Now merged cells have appropriate merge flags.
        // ExEnd:ConvertToHorizontallyMergedCells
        System.out.println("\nNow merged cells have appropriate merge flags.\nFile saved at " + dataDir);
    }

}
