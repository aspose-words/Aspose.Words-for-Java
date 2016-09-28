package com.aspose.words.examples.programming_documents.find_replace;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

import java.util.regex.Pattern;

public class ReplaceWithRegex {
	
	private static final String dataDir = Utils.getSharedDataDir(ReplaceWithRegex.class) + "FindAndReplace/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document(dataDir + "ReplaceWithRegex.doc");
		doc.getRange().replace(Pattern.compile("[s|m]ad"), "happy");
		doc.save(dataDir + "ReplaceWithRegex_Out.doc");
	}
}