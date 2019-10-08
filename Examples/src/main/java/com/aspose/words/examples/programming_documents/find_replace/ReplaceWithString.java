package com.aspose.words.examples.programming_documents.find_replace;

import com.aspose.words.Document;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.examples.Utils;

public class ReplaceWithString {

    public static final String dataDir = Utils.getSharedDataDir(ReplaceWithString.class) + "FindAndReplace/";

    public static void main(String[] args) throws Exception {

        //ExStart:ReplaceWithString
        Document doc = new Document(dataDir + "ReplaceWithString.doc");
        doc.getRange().replace("sad", "bad", new FindReplaceOptions(FindReplaceDirection.FORWARD));
        doc.save(dataDir + "ReplaceWithString_out.doc");
        //ExEnd:ReplaceWithString
    }
}