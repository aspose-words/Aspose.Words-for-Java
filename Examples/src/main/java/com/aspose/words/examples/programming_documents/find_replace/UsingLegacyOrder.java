package com.aspose.words.examples.programming_documents.find_replace;

import java.util.regex.Pattern;

import com.aspose.words.Document;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.examples.Utils;

public class UsingLegacyOrder {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(UsingLegacyOrder.class);

        FineReplaceUsingLegacyOrder(dataDir);
	}
	
	// ExStart:FineReplaceUsingLegacyOrder
    public static void FineReplaceUsingLegacyOrder(String dataDir) throws Exception
    {
        // Open the document.
        Document doc = new Document(dataDir + "source.docx");
        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new ReplacingCallback());
        options.setUseLegacyOrder(true);

        doc.getRange().replace(Pattern.compile("\\[(.*?)\\]"), "", options);

        dataDir = dataDir + "usingLegacyOrder_out.doc";
        doc.save(dataDir);
    }

    private static class ReplacingCallback implements IReplacingCallback
    {
        public int replacing(ReplacingArgs args) {
        	System.out.print(args.getMatch().group());
            return ReplaceAction.REPLACE;
        }
    }
    // ExEnd:FineReplaceUsingLegacyOrder
}