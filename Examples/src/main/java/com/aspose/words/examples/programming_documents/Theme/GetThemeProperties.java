/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.Theme;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class GetThemeProperties
{
    public static void main(String[] args) throws Exception
    {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(GetThemeProperties.class);

		Document doc = new Document(dataDir +"Document.doc");

		Theme theme = doc.getTheme();
		// Major (Headings) font for Latin characters.
		System.out.println(theme.getMajorFonts().getLatin());
		// Minor (Body) font for EastAsian characters.
		System.out.println(theme.getMinorFonts().getEastAsian());
		// Color for theme color Accent 1.
		System.out.println(theme.getColors().getAccent1());
		// Save the document to disk.
		//doc.save(dataDir + "TestFile.AutoFitToContents Out.doc");

		//System.out.println("Table auto fit to contents successfully.");
    }
}