package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.aspose.words.examples.Utils;

public class SpecifyDefaultFontToUseWhenRendering {
	
	private static final String dataDir = Utils.getSharedDataDir(SpecifyDefaultFontToUseWhenRendering.class) + "RenderingAndPrinting/";
	
	public static void main(String[] args) throws Exception {
		
		Document doc = new Document(dataDir + "Rendering.doc");

		// If the default font defined here cannot be found during rendering then the closest font on the machine is used instead.
		FontSettings.getDefaultInstance().setDefaultFontName("Arial Unicode MS");

		// Now the set default font is used in place of any missing fonts during any rendering calls.
		doc.save(dataDir + "Rendering.SetDefaultFont Out.pdf");
		doc.save(dataDir + "Rendering.SetDefaultFont Out.xps");
	}

}
