package com.aspose.words.examples.programming_documents.find_replace;

import com.aspose.words.Document;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.examples.Utils;

import java.util.regex.Pattern;

//ExStart:1
public class ReplaceWithEvaluator {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ReplaceWithEvaluator.class);

        Document doc = new Document(dataDir + "Document.doc");
        doc.getRange().replace(Pattern.compile("[s|m]ad"), new ReplaceCallback(), true);

        doc.save(dataDir + "output.doc");
    }
}

class ReplaceCallback implements IReplacingCallback {
    private int count = 0;

    @Override
    public int replacing(ReplacingArgs args) throws Exception {
        count++;
        args.setReplacement("HAPPY-" + count);
        return ReplaceAction.REPLACE;
    }
}
//ExEnd:1

