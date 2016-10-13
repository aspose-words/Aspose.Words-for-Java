
package com.aspose.words.examples.programming_documents.Theme;

import com.aspose.words.Document;
import com.aspose.words.Theme;
import com.aspose.words.examples.Utils;
import javafx.scene.paint.Color;


public class SetThemeProperties
{
    public static void main(String[] args) throws Exception
    {

		// The path to the documents directory.
		String dataDir = Utils.getDataDir(SetThemeProperties.class);

		Document doc = new Document(dataDir +"Document.doc");

		Theme theme = doc.getTheme();
		// Set Times New Roman font as Body theme font for Latin Character.
		theme.getMinorFonts().setLatin("Algerian");
		// Set Color.Gold for theme color Hyperlink.
		theme.getColors().setHyperlink(java.awt.Color.DARK_GRAY);
		doc.save(dataDir+  "output.doc");

	}
}