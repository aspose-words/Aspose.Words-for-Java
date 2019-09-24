package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;

public class AdvancedMailMergeFeatures {

    public static void main(String[] args) {

    }

    /**
     * Add a mapping when a merge field in a document and a data field in a data
     * source have different names.
     */
    private static void addMappingWhenMergeFieldAndDataFieldHaveDifferentNames(Document doc) {
        //ExStart:addMappingWhenMergeFieldAndDataFieldHaveDifferentNames
        doc.getMailMerge().getMappedDataFields().add("MyFieldName_InDocument", "MyFieldName_InDataSource");
        //ExEnd:addMappingWhenMergeFieldAndDataFieldHaveDifferentNames
    }

    /**
     * Get names of all merge fields in a document.
     */
    private static void getNamesOfAllMergeFields(Document doc) throws Exception {
        //ExStart:getNamesOfAllMergeFields
        String[] fieldNames = doc.getMailMerge().getFieldNames();
        //ExEnd:getNamesOfAllMergeFields
    }

    /**
     * Delete all merge fields from a document without executing mail merge.
     */
    private static void deletingMergeFields(Document doc) throws Exception {
        //ExStart:deletingMergeFields
        doc.getMailMerge().deleteFields();
        //ExEnd:deletingMergeFields
    }
}