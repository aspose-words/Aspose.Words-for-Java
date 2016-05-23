/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.examples.Utils;

import java.util.regex.Pattern;

//TODO: re-check output
public class ReplaceWithEvaluator
{
    public static void main(String[] args) throws Exception
    {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ReplaceWithEvaluator.class);

        Document doc = new Document(dataDir + "Document.doc");
       /// doc.getRange().replace(Pattern.compile("[s|m]ad"));
        doc.getRange().replace(Pattern.compile("[s|m]ad]"), new ReplaceCallback(), true);

        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }

}

class ReplaceCallback implements IReplacingCallback {
    @Override
    public int replacing(ReplacingArgs replacingArgs) throws Exception {
        return 1;
    }
}

