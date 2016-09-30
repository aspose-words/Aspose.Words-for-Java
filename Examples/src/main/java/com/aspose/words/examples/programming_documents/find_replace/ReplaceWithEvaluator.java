package com.aspose.words.examples.programming_documents.find_replace;

import java.util.regex.Pattern;

import com.aspose.words.Document;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.examples.Utils;

public class ReplaceWithEvaluator {

	private static final String dataDir = Utils.getSharedDataDir(ReplaceWithEvaluator.class) + "FindAndReplace/";

	public static void main(String[] args) throws Exception {
		Document doc = new Document(dataDir + "Range.ReplaceWithEvaluator.doc");

		FindReplaceOptions options = new FindReplaceOptions();
		options.ReplacingCallback = new MyReplaceEvaluator();

		doc.getRange().replace(Pattern.compile("[s|m]ad"), "", options);
		doc.save(dataDir + "Range.ReplaceWithEvaluator_Out.doc");
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
