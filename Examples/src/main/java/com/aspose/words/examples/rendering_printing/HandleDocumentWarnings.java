package com.aspose.words.examples.rendering_printing;

import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningInfo;
import com.aspose.words.WarningInfoCollection;
import com.aspose.words.WarningType;

//ExStart:HandleDocumentWarnings
public class HandleDocumentWarnings implements IWarningCallback {
	/**
	 * Our callback only needs to implement the "Warning" method. This method is
	 * called whenever there is a potential issue during document processing. The
	 * callback can be set to listen for warnings generated during document load
	 * and/or document save.
	 */

	public void warning(WarningInfo info) {
		// For now type of warnings about unsupported metafile records changed from
		// DataLoss/UnexpectedContent to MinorFormattingLoss.
		if (info.getWarningType() == WarningType.MINOR_FORMATTING_LOSS) {
			System.out.println("Unsupported operation: " + info.getDescription());
			mWarnings.warning(info);
		}
	}

	public WarningInfoCollection mWarnings = new WarningInfoCollection();
}
//ExEnd:HandleDocumentWarnings