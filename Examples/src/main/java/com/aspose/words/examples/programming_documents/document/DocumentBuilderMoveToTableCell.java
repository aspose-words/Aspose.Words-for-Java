package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;
import org.testng.Assert;

public class DocumentBuilderMoveToTableCell {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderMoveToTableCell
        String dataDir = Utils.getDataDir(DocumentBuilderMoveToTableCell.class);

        Document doc = new Document(dataDir + "Tables.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder to row 3, cell 4 of the first table.
        builder.moveToCell(0, 2, 3, 0);
        builder.write("\nCell contents added by DocumentBuilder");
        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(table.getRows().get(2).getCells().get(3), builder.getCurrentNode().getParentNode().getParentNode());
        Assert.assertEquals("Cell contents added by DocumentBuilderCell 3 contents\u0007", table.getRows().get(2).getCells().get(3).getText().trim());
        //ExEnd:DocumentBuilderMoveToTableCell

    }
}