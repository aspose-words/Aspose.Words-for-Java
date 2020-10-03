package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldNextIf;
import com.aspose.words.FieldSkipIf;
import com.aspose.words.FieldType;

public class AdvancedMailMergeFeatures {

    public static void main(String[] args) throws Exception{

    	CompareTwoExpressions();
    }

    private static void CompareTwoExpressions() throws Exception {
    	Document doc = new Document();
    	DocumentBuilder builder = new DocumentBuilder(doc);
    	
    	//ExStart: CompareTwoExpressions
    	// Use NextIf field
        FieldNextIf fieldNextIf = (FieldNextIf)builder.insertField(FieldType.FIELD_NEXT_IF, true);

        // Or use SkipIf field
        FieldSkipIf fieldSkipIf = (FieldSkipIf)builder.insertField(FieldType.FIELD_SKIP_IF, true);

        // Compare two expressions
        fieldNextIf.setLeftExpression("3");
        fieldNextIf.setRightExpression("1 + 2");
        fieldNextIf.setComparisonOperator("=");
    	//ExEnd: CompareTwoExpressions
    }
    /**
     * Add a mapping when a merge field in a document and a data field in a data
     * source have different names.
     */
    private static void addMappingWhenMergeFieldAndDataFieldHaveDifferentNames(Document doc) {
        doc.getMailMerge().getMappedDataFields().add("MyFieldName_InDocument", "MyFieldName_InDataSource");
    }

    /**
     * Get names of all merge fields in a document.
     */
    private static void getNamesOfAllMergeFields(Document doc) throws Exception {
        String[] fieldNames = doc.getMailMerge().getFieldNames();
    }

    /**
     * Delete all merge fields from a document without executing mail merge.
     */
    private static void deletingMergeFields(Document doc) throws Exception {
        doc.getMailMerge().deleteFields();
    }
}