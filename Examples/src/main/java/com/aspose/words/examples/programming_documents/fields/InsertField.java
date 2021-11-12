package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataTable;
import org.testng.Assert;

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
        fieldIf(dataDir);
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

    private static void fieldIf(String dataDir) throws Exception {
        System.out.println("==== fieldIf ====");
        //ExStart:fieldIf
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Statement 1: ");

        // Use document builder to insert an if field
        FieldIf fieldIf = (FieldIf) builder.insertField(FieldType.FIELD_IF, true);

        // The if field will output either the TrueText or FalseText string into the document, depending on the truth of the statement
        // In this case, "0 = 1" is incorrect, so the output will be "False"
        fieldIf.setLeftExpression("0");
        fieldIf.setComparisonOperator("=");
        fieldIf.setRightExpression("1");
        fieldIf.setTrueText("True");
        fieldIf.setFalseText("False");

        System.out.println(" IF  0 = 1 True False".equals(fieldIf.getFieldCode()));
        System.out.println(FieldIfComparisonResult.getName(fieldIf.evaluateCondition()));

        // This time, the statement is correct, so the output will be "True"
        builder.write("\nStatement 2: ");
        fieldIf = (FieldIf) builder.insertField(FieldType.FIELD_IF, true);
        fieldIf.setLeftExpression("5");
        fieldIf.setComparisonOperator("=");
        fieldIf.setRightExpression("2 + 3");
        fieldIf.setTrueText("True");
        fieldIf.setFalseText("False");

        System.out.println(" IF  5 = \"2 + 3\" True False".equals(fieldIf.getFieldCode()));
        System.out.println(FieldIfComparisonResult.getName(fieldIf.evaluateCondition()));

        doc.updateFields();
        //ExEnd:fieldIf
        doc.save(dataDir + "Field.If.docx");
    }

    public void fieldAdvance(String dataDir) throws Exception {
        //ExStart:fieldAdvance
        Document doc = new Document(dataDir + "in.doc");

        // Get paragraph you want to append this merge field to
        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(1);

        // Create instance of FieldAdvance class and lets build the above field code
        FieldAdvance field = (FieldAdvance) para.appendField(FieldType.FIELD_ADVANCE, false);
        field.setRightOffset("5");
        field.setUpOffset("5");

        field.update();

        doc.save(dataDir + "output.docx");
        //ExEnd:fieldAdvance
    }

    public void fieldAsk(String dataDir) throws Exception {
        //ExStart:fieldAsk
        Document doc = new Document(dataDir + "in.doc");

        // Get paragraph you want to append this merge field to
        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(1);

        // Create instance of FieldAsk class and lets build the above field code
        FieldAsk field = (FieldAsk) para.appendField(FieldType.FIELD_ASK, false);
        field.setPromptText("Please provide a response for this ASK field");
        field.setDefaultResponse("Response from within the field.");

        field.update();

        doc.save(dataDir + "output.docx");
        //ExEnd:fieldAsk
    }

    public void fieldIncludeText(String dataDir) throws Exception {
        //ExStart:fieldIncludeText
        Document doc = new Document(dataDir + "in.doc");

        // Get paragraph you want to append this merge field to
        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(1);

        // Create instance of FieldIncludeText class and lets build the above field code
        FieldIncludeText field = (FieldIncludeText) para.appendField(FieldType.FIELD_INCLUDE_TEXT, false);
        field.setBookmarkName("bookmark");
        field.setSourceFullName(dataDir + "IncludeText.docx");

        field.update();

        doc.save(dataDir + "output.docx");
        //ExEnd:fieldIncludeText
    }
}




