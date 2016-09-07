package com.aspose.words.examples.document_object_model;

import com.aspose.words.Document;

public class GetNodeType {

	public static void main(String[] args) throws Exception {
		Document doc = new Document();
		// Returns NodeType.Document
		int type = doc.getNodeType();
	}
}
