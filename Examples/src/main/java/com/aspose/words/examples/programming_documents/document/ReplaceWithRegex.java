package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

import java.util.regex.Pattern;

public class ReplaceWithRegex {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ReplaceWithRegex.class);

		Document doc = new Document(dataDir + "Document.doc");
		doc.getRange().replace(Pattern.compile("[s|m]ad"), "happy");
		doc.save(dataDir + "output.doc");
	}

}