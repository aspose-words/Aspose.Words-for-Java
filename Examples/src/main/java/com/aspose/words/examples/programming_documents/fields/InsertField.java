package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldCompare;
import com.aspose.words.FieldType;
import com.aspose.words.examples.Utils;

public class InsertField {
    public static void main(String[] args) throws Exception {
        //ExStart:InsertField
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertField.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD MyFieldName \\* MERGEFORMAT");

        doc.save(dataDir + "output.docx");
        //ExEnd:InsertField

        fieldCompare(dataDir);
    }

    private static void fieldCompare(String dataDir) throws Exception {
        //ExStart:fieldCompare
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a compare field using a document builder
        FieldCompare field = (FieldCompare) builder.insertField(FieldType.FIELD_COMPARE, true);

        // Construct a comparison statement
        field.setLeftExpression("3");
        field.setComparisonOperator("<");
        field.setRightExpression("2");

        // The compare field will print a "0" or "1" depending on the truth of its statement
        // The result of this statement is false, so a "0" will be shown up in the document
        System.out.println(" COMPARE  3 < 2".equals(field.getFieldCode()));

        builder.writeln();

        // Here a "1" will show up, because the statement is true
        field = (FieldCompare) builder.insertField(FieldType.FIELD_COMPARE, true);
        field.setLeftExpression("5");
        field.setComparisonOperator("=");
        field.setRightExpression("2 + 3");

        System.out.println(" COMPARE  5 = \"2 + 3\"".equals(field.getFieldCode()));

        doc.updateFields();
        //ExEnd:fieldCompare
        doc.save(dataDir + "Field.Compare.docx");
    }
}




