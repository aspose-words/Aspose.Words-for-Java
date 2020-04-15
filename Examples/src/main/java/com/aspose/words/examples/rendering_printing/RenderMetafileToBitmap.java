package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.MetafileRenderingMode;
import com.aspose.words.MetafileRenderingOptions;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;

public class RenderMetafileToBitmap {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(RenderMetafileToBitmap.class);

		// ExStart:RenderMetafileToBitmap
		// Load the document from disk.
		Document doc = new Document(dataDir + "Rendering.doc");

		MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
		metafileRenderingOptions.setEmulateRasterOperations(false);
		metafileRenderingOptions.setRenderingMode(MetafileRenderingMode.VECTOR_WITH_FALLBACK);

		// If Aspose.Words cannot correctly render some of the metafile records to
		// vector graphics then Aspose.Words renders this metafile to a bitmap.
		HandleDocumentWarnings callback = new HandleDocumentWarnings();
		doc.setWarningCallback(callback);

		PdfSaveOptions saveOptions = new PdfSaveOptions();
		saveOptions.setMetafileRenderingOptions(metafileRenderingOptions);

		doc.save(dataDir + "PdfSaveOptions.HandleRasterWarnings_out.pdf", saveOptions);
		// ExEnd:RenderMetafileToBitmap
	}

}
