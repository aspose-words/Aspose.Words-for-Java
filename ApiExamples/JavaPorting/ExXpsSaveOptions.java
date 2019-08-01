// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.XpsSaveOptions;


@Test
public class ExXpsSaveOptions extends ApiExampleBase
{
    @Test
    public void optimizeOutput() throws Exception
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.OptimizeOutput
        //ExSummary:Shows how to optimize document objects while saving to xps.
        Document doc = new Document(getMyDir() + "XPSOutputOptimize.docx");

        XpsSaveOptions saveOptions = new XpsSaveOptions(); { saveOptions.setOptimizeOutput(true); }

        doc.save(getArtifactsDir() + "XPSOutputOptimize.xps", saveOptions);
        //ExEnd
    }
}
