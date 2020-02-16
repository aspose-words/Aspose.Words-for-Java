package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Document;
import com.aspose.words.XpsSaveOptions;
import org.testng.annotations.Test;

@Test
public class ExXpsSaveOptions extends ApiExampleBase {
    @Test
    public void optimizeOutput() throws Exception {
        //ExStart
        //ExFor:FixedPageSaveOptions.OptimizeOutput
        //ExSummary:Shows how to optimize document objects while saving to xps.
        Document doc = new Document(getMyDir() + "Unoptimized document.docx");

        XpsSaveOptions saveOptions = new XpsSaveOptions();
        saveOptions.setOptimizeOutput(true);

        doc.save(getArtifactsDir() + "XpsSaveOptions.OptimizeOutputF.xps", saveOptions);
        //ExEnd
    }
}
