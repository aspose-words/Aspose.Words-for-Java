package com.aspose.words.examples.asposefeatures.tables.repeatheaderrow;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class AsposeRepeatHeaderRow
{
    public static void main(String[] args) throws Exception
    {
        try {
            // The path to the documents directory.
            String dataDir = Utils.getDataDir(AsposeRepeatHeaderRow.class);

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.startTable();
            builder.getRowFormat().setHeadingFormat(true);
            builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
            builder.getCellFormat().setWidth(100);
            builder.insertCell();
            builder.writeln("Heading row 1");
            builder.endRow();
            builder.insertCell();
            builder.writeln("Heading row 2");
            builder.endRow();

            builder.getCellFormat().setWidth(50);
            builder.getParagraphFormat().clearFormatting();

            // Insert some content so the table is long enough to continue onto the next page.
            for (int i = 0; i < 50; i++)
            {
                builder.insertCell();
                builder.getRowFormat().setHeadingFormat(false);
                builder.write("Column 1 Text");
                builder.insertCell();
                builder.write("Column 2 Text");
                builder.endRow();
            }

            doc.save(dataDir + "RepearHeaderRows_Out.doc");

            System.out.println("Done.");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}
