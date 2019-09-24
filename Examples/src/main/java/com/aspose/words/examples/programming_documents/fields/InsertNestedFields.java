package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class InsertNestedFields {
    public static void main(String[] args) throws Exception {

        //ExStart:InsertNestedFields
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertNestedFields.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert few page breaks (just for testing)
        for (int i = 0; i < 5; i++)
            builder.insertBreak(BreakType.PAGE_BREAK);

        // Move DocumentBuilder cursor into the primary footer.
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);

        // We want to insert a field like this:
        // { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
        Field field = builder.insertField("IF ");
        builder.moveTo(field.getSeparator());
        builder.insertField("PAGE");
        builder.write(" <> ");
        builder.insertField("NUMPAGES");
        builder.write(" \"See Next Page\" \"Last Page\" ");

        // Finally update the outer field to recalcaluate the final value. Doing this will automatically update
        // the inner fields at the same time.
        field.update();

        doc.save(dataDir + "output.docx");
        //ExEnd:InsertNestedFields


    }
}




