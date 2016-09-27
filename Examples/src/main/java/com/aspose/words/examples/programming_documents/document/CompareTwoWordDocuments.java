package com.aspose.words.examples.programming_documents.document;

import java.util.Date;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.tables.creation.BuildTableFromDataTable;

public class CompareTwoWordDocuments {
	
	private static final String dataDir = Utils.getSharedDataDir(BuildTableFromDataTable.class) + "Document/";
	
	public static void main(String[] args) throws Exception {
		// Example Shows Normal Comparison Case
		normalComparisonCase();
		
		// Case when Document has Revisions already so Comparison is not Possible
		caseWhenDocumentHasRevisions();
		
		// Shows how to test that Word Documents are "Equal"
		wordDocumentsAreEqual();
	}

	public static void normalComparisonCase() throws Exception {
		Document docA = new Document(dataDir + "DocumentA.doc");
		Document docB = new Document(dataDir + "DocumentB.doc");
		docA.compare(docB, "user", new Date()); // docA now contains changes as revisions
	}
	
	public static void caseWhenDocumentHasRevisions() throws Exception {
		Document docA = new Document(dataDir + "DocumentA.doc");
		Document docB = new Document(dataDir + "DocumentB.doc");
		docA.compare(docB, "user", new Date());    // exception is thrown.
	}
	
	public static void wordDocumentsAreEqual() throws Exception {
		Document docA = new Document(dataDir + "DocumentA.doc");
		Document docB = new Document(dataDir + "DocumentB.doc");
		docA.compare(docB, "user", new Date());
		if(docA.getRevisions().getCount() == 0)
		    System.out.println("Documents are equal");
	}
}
