package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.CompareOptions;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Granularity;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.tables.creation.BuildTableFromDataTable;

import java.util.Date;

public class CompareTwoWordDocuments {

	private static final String dataDir = Utils.getSharedDataDir(BuildTableFromDataTable.class) + "Document/";

	public static void main(String[] args) throws Exception {
		// ExStart:CompareTwoWordDocuments
		// Example Shows Normal Comparison Case
		normalComparisonCase();

		// Case when Document has Revisions already so Comparison is not Possible
		caseWhenDocumentHasRevisions();

		// Shows how to test that Word Documents are "Equal"
		wordDocumentsAreEqual();

		SpecifyComparisonGranularity(dataDir);
		// ExEnd:CompareTwoWordDocuments
	}

	public static void normalComparisonCase() throws Exception {
		// ExStart:normalComparisonCase
		Document docA = new Document(dataDir + "DocumentA.doc");
		Document docB = new Document(dataDir + "DocumentB.doc");
		docA.compare(docB, "user", new Date()); // docA now contains changes as revisions
		// ExEnd:normalComparisonCase
	}

	public static void caseWhenDocumentHasRevisions() throws Exception {
		// ExStart:caseWhenDocumentHasRevisions
		Document docA = new Document(dataDir + "DocumentA.doc");
		Document docB = new Document(dataDir + "DocumentB.doc");
		docA.compare(docB, "user", new Date()); // exception is thrown.
		// ExEnd:caseWhenDocumentHasRevisions
	}

	public static void wordDocumentsAreEqual() throws Exception {
		// ExStart:wordDocumentsAreEqual
		Document docA = new Document(dataDir + "DocumentA.doc");
		Document docB = new Document(dataDir + "DocumentB.doc");
		docA.compare(docB, "user", new Date());
		if (docA.getRevisions().getCount() == 0)
			System.out.println("Documents are equal");
		// ExEnd:wordDocumentsAreEqual
	}

	public static void SpecifyComparisonGranularity(String dataDir) throws Exception {
		// ExStart:SpecifyComparisonGranularity
		DocumentBuilder builderA = new DocumentBuilder(new Document());
		DocumentBuilder builderB = new DocumentBuilder(new Document());

		builderA.writeln("This is A simple word");
		builderB.writeln("This is B simple words");

		CompareOptions co = new CompareOptions();
		co.setGranularity(Granularity.CHAR_LEVEL);

		builderA.getDocument().compare(builderB.getDocument(), "author", new Date(), co);
		// ExEnd:SpecifyComparisonGranularity
	}
}
