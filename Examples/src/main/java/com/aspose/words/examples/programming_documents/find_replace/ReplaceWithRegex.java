package com.aspose.words.examples.programming_documents.find_replace;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.examples.Utils;

import java.util.regex.Pattern;

public class ReplaceWithRegex {

	private static final String dataDir = Utils.getSharedDataDir(ReplaceWithRegex.class) + "FindAndReplace/";

	public static void main(String[] args) throws Exception {
		// ExStart:ReplaceWithRegex
		Document doc = new Document(dataDir + "ReplaceWithRegex.doc");
		FindReplaceOptions options = new FindReplaceOptions();
		doc.getRange().replace(Pattern.compile("[s|m]ad"), "happy", options);
		doc.save(dataDir + "ReplaceWithRegex_Out.doc");
		// ExEnd:ReplaceWithRegex

		RecognizeAndSubstitutionsWithinReplacementPatterns(dataDir);
	}

	public static void RecognizeAndSubstitutionsWithinReplacementPatterns(String dataDir) throws Exception {
		// ExStart:RecognizeAndSubstitutionsWithinReplacementPatterns
		// Create new document.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// Write some text.
		builder.write("Jason give money to Paul.");

		// Replace text using substitutions.
		FindReplaceOptions options = new FindReplaceOptions();
		options.setUseSubstitutions(true);
		doc.getRange().replace(Pattern.compile("([A-z]+) give money to ([A-z]+)"), "$2 take money from $1", options);
		// ExEnd:RecognizeAndSubstitutionsWithinReplacementPatterns
		System.out.println(doc.getText()); // The output is: Paul take money from Jason.\f
	}
}