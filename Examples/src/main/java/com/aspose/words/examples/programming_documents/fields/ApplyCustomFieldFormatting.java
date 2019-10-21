package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Field;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 5/29/2017.
 */
public class ApplyCustomFieldFormatting {
    public static void main(String[] args) throws Exception {
        //ExStart:ApplyCustomFieldFormatting
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApplyCustomFieldFormatting.class);
        DocumentBuilder builder = new DocumentBuilder();
        Document document = builder.getDocument();

        Field field = builder.insertField("=-1234567.89 \\# \"### ### ###.000\"", null);
        document.getFieldOptions().setResultFormatter(new FieldResultFormatter("[%0$s]", null));

        field.update();
        document.save(dataDir + "FormatFieldResult_out.docx");
        //ExEnd:ApplyCustomFieldFormatting
    }


}