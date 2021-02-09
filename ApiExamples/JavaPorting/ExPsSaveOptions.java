// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.PsSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.Section;
import com.aspose.words.MultiplePagesType;
import org.testng.annotations.DataProvider;


@Test
public class ExPsSaveOptions extends ApiExampleBase
{
    @Test (dataProvider = "useBookFoldPrintingSettingsDataProvider")
    public void useBookFoldPrintingSettings(boolean renderTextAsBookFold) throws Exception
    {
        //ExStart
        //ExFor:PsSaveOptions
        //ExFor:PsSaveOptions.SaveFormat
        //ExFor:PsSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the Postscript format in the form of a book fold.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Create a "PsSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to PostScript.
        // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
        // in the output Postscript document in a way that helps us make a booklet out of it.
        // Set the "UseBookFoldPrintingSettings" property to "false" to save the document normally.
        PsSaveOptions saveOptions = new PsSaveOptions();
        {
            saveOptions.setSaveFormat(SaveFormat.PS);
            saveOptions.setUseBookFoldPrintingSettings(renderTextAsBookFold);
        }

        // If we are rendering the document as a booklet, we must set the "MultiplePages"
        // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
        for (Section s : (Iterable<Section>) doc.getSections())
        {
            s.getPageSetup().setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);
        }

        // Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
        // and the contents will line up in a way that creates a booklet.
        doc.save(getArtifactsDir() + "PsSaveOptions.UseBookFoldPrintingSettings.ps", saveOptions);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "useBookFoldPrintingSettingsDataProvider")
	public static Object[][] useBookFoldPrintingSettingsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}
}
