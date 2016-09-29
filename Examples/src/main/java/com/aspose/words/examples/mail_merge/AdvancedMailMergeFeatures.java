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