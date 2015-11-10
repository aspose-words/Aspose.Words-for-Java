package com.aspose.words.examples.asposefeatures.workingwithtext.findnreplacetxt;

import java.util.regex.Pattern;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

// For more info please visit http://www.aspose.com/docs/display/wordsjava/Find+and+Replace+Overview
public class AsposeFindnReplace
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeFindnReplace.class);

        Document doc = new Document(dataDir + "replaceDoc.doc");

        // Replaces all 'sad' and 'mad' occurrences with 'bad'
        doc.getRange().replace("sad", "bad", false, true); 

        // Replaces all 'sad' and 'mad' occurrences with 'bad'
        doc.getRange().replace(Pattern.compile("[s|m]ad"), "bad");

        doc.save(dataDir + "AsposeReplaced.doc", SaveFormat.DOC);

        System.out.println("Done.");
    }
}