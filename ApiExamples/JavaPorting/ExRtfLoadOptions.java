// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.System.msString;
import org.testng.annotations.DataProvider;


@Test
public class ExRtfLoadOptions extends ApiExampleBase
{
    @Test (dataProvider = "recognizeUtf8TextDataProvider")
    public void recognizeUtf8Text(boolean doRecognizeUtb8Text) throws Exception
    {
        //ExStart
        //ExFor:RtfLoadOptions
        //ExFor:RtfLoadOptions.#ctor
        //ExFor:RtfLoadOptions.RecognizeUtf8Text
        //ExSummary:Shows how to detect UTF8 characters during import.
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        {
            loadOptions.setRecognizeUtf8Text(doRecognizeUtb8Text);
        }

        Document doc = new Document(getMyDir() + "UTF-8 characters.rtf", loadOptions);

        Assert.assertEquals(
            doRecognizeUtb8Text
                ? "“John Doe´s list of currency symbols”™\r€, ¢, £, ¥, ¤"
                : "â€œJohn DoeÂ´s list of currency symbolsâ€\u009dâ„¢\râ‚¬, Â¢, Â£, Â¥, Â¤",
            msString.trim(doc.getFirstSection().getBody().getText()));
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
