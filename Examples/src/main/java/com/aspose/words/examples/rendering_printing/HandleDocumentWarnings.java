package com.aspose.words.examples.rendering_printing;

import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningInfo;
import com.aspose.words.WarningType;

public class HandleDocumentWarnings implements IWarningCallback {
	/**
	 * Our callback only needs to implement the "Warning" method. This method is
	 * called whenever there is a potential issue during document processing.
	 * The callback can be set to listen for warnings generated during document
	 * load and/or document save.
	 */
	public void warning(WarningInfo info) {
		// We are only interested in fonts being substituted.
		if (info.getWarningType() == WarningType.FONT_SUBSTITUTION) {
			System.out.println("Font substitution: " + info.getDescription());
		}
	}
}