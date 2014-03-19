/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

package programmingwithdocuments.workingwithfields.insertnestedfields.java;

import com.aspose.words.*;

public class InsertNestedFields
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmingwithdocuments/workingwithfields/insertnestedfields/data/";

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

        doc.save(dataDir + "InsertNestedFields Out.docx");
    }
}




