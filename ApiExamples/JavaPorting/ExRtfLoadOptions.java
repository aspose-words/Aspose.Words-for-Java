// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.RtfLoadOptions;
import com.aspose.words.Document;
import org.testng.Assert;
import org.testng.annotations.DataProvider;


@Test
public class ExRtfLoadOptions extends ApiExampleBase
{
    @Test (dataProvider = "recognizeUtf8TextDataProvider")
    public void recognizeUtf8Text(boolean recognizeUtf8Text) throws Exception
    {
        //ExStart
        //ExFor:RtfLoadOptions
        //ExFor:RtfLoadOptions.#ctor
        //ExFor:RtfLoadOptions.RecognizeUtf8Text
        //ExSummary:Shows how to detect UTF-8 characters while loading an RTF document.
        // Create an "RtfLoadOptions" object to modify how we load an RTF document.
        RtfLoadOptions loadOptions = new RtfLoadOptions();

        // Set the "RecognizeUtf8Text" property to "false" to assume that the document uses the ISO 8859-1 charset
        // and loads every character in the document.
        // Set the "RecognizeUtf8Text" property to "true" to parse any variable-length characters that may occur in the text.
        loadOptions.setRecognizeUtf8Text(recognizeUtf8Text);

        Document doc = new Document(getMyDir() + "UTF-8 characters.rtf", loadOptions);

        Assert.assertEquals(
            recognizeUtf8Text
                ? "“John Doe´s list of currency symbols”™\r" +
                  "€, ¢, £, ¥, ¤"
                : "â€œJohn DoeÂ´s list of currency symbolsâ€\u009dâ„¢\r" +
                  "â‚¬, Â¢, Â£, Â¥, Â¤",
            doc.getFirstSection().getBody().getText().trim());
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "recognizeUtf8TextDataProvider")
	public static Object[][] recognizeUtf8TextDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}
}
