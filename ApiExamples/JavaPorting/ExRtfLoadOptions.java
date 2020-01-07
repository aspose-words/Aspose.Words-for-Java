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


@Test
public class ExRtfLoadOptions extends ApiExampleBase
{
    @Test
    public void recognizeUtf8Text() throws Exception
    {
        //ExStart
        //ExFor:RtfLoadOptions
        //ExFor:RtfLoadOptions.#ctor
        //ExFor:RtfLoadOptions.RecognizeUtf8Text
        //ExSummary:Shows how to detect UTF8 characters during import.
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        {
            loadOptions.setRecognizeUtf8Text(true);
        }

        Document doc = new Document(getMyDir() + "RtfLoadOptions.RecognizeUtf8Text.rtf", loadOptions);
        doc.save(getArtifactsDir() + "RtfLoadOptions.RecognizeUtf8Text.rtf");
        //ExEnd
    }
}
