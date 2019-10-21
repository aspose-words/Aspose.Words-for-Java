package com.aspose.words.examples.programming_documents.find_replace;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.util.regex.Pattern;

public class ReplaceWithEvaluator {

    private static final String dataDir = Utils.getSharedDataDir(ReplaceWithEvaluator.class) + "FindAndReplace/";

    public static void main(String[] args) throws Exception {
        //ExStart:ReplaceWithEvaluator
        Document doc = new Document(dataDir + "Range.ReplaceWithEvaluator.doc");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new MyReplaceEvaluator());

        doc.getRange().replace(Pattern.compile("[s|m]ad"), "", options);
        doc.save(dataDir + "Range.ReplaceWithEvaluator_Out.doc");
        //ExEnd:ReplaceWithEvaluator
    }
}

//ExStart:MyReplaceEvaluator
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
//ExEnd:MyReplaceEvaluator
}
