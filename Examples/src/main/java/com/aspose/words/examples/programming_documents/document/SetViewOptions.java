package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.ViewType;
import com.aspose.words.examples.Utils;

public class SetViewOptions {

	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(SetViewOptions.class) + "Document/";
		// ExStart:SetViewOptions
		Document doc = new Document(dataDir + "Document.doc");
		doc.getViewOptions().setViewType(ViewType.PAGE_LAYOUT);
		doc.getViewOptions().setZoomPercent(50);
		doc.save(dataDir + "Document.SetZoom_out.doc");
		// ExEnd:SetViewOptions
	}
}
