package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 5/29/2017.
 */
public class ApplyCellPadding {

    public static void main(String[] args) throws Exception {

        //ExStart:ApplyCellPadding
        String dataDir = Utils.getDataDir(ApplyCellPadding.class);
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();

        //Sets the amount of space (in points) to add to the left/top/right/bottom of the contents of cell.
        builder.getCellFormat().setPaddings(30, 50, 30, 50);
        builder.writeln("I'm a wonderful formatted cell.");

        builder.endRow();
        builder.endTable();

        dataDir = dataDir + "Table.SetCellPadding_out.doc";

        //Save the document to disk.
        doc.save(dataDir);
        //ExEnd:ApplyCellPadding

    }
}
