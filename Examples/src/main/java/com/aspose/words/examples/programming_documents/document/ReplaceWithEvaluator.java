package com.aspose.words.examples.programming_documents.document;

import java.util.regex.Pattern;

import com.aspose.words.Document;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.examples.Utils;

public class ReplaceWithEvaluator {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ReplaceWithEvaluator.class);

		Document doc = new Document(dataDir + "Document.doc");
		doc.getRange().replace(Pattern.compile("[s|m]ad"), new MyReplaceEvaluator(), true);
		doc.save(dataDir + "output.doc");
	}
}

class MyReplaceEvaluator implements IReplacingCallback {
	private int mMatchNumber;

	/**
	 * This is called during a replace operation each time a match is found.
	 * This method appends a number to the match string and returns it as a
	 * replacement string.
	 */
	public int replacing(ReplacingArgs e) throws Exception {
		e.setReplacement(e.getMatch().group() + Integer.toString(mMatchNumber));
		mMatchNumber++;
		return ReplaceAction.REPLACE;
	}

}
