// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.System.IO.FileInfo;
import org.testng.Assert;
import org.testng.annotations.DataProvider;


@Test
public class ExXpsSaveOptions extends ApiExampleBase
{
    @Test (dataProvider = "optimizeOutputDataProvider")
    public void optimizeOutput(boolean optimizeOutput) throws Exception
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.OptimizeOutput
        //ExSummary:Shows how to optimize document objects while saving to xps.
        Document doc = new Document(getMyDir() + "Unoptimized document.docx");

        // When saving to .xps, we can use SaveOptions to optimize the output in some cases
        XpsSaveOptions saveOptions = new XpsSaveOptions(); { saveOptions.setOptimizeOutput(optimizeOutput); }

        doc.save(getArtifactsDir() + "XpsSaveOptions.OptimizeOutput.xps", saveOptions);

        // The input document had adjacent runs with the same formatting, which, if output optimization was enabled,
        // have been combined to save space
        FileInfo outFileInfo = new FileInfo(getArtifactsDir() + "XpsSaveOptions.OptimizeOutput.xps");

        if (optimizeOutput)
            Assert.assertTrue(outFileInfo.getLength() < 45000);
        else
            Assert.assertTrue(outFileInfo.getLength() > 60000);
        //ExEnd

        TestUtil.docPackageFileContainsString(
            optimizeOutput
                ? "Glyphs OriginX=\"34.294998169\" OriginY=\"10.31799984\" " +
                  "UnicodeString=\"This document contains complex content which can be optimized to save space when \""
                : "<Glyphs OriginX=\"34.294998169\" OriginY=\"10.31799984\" UnicodeString=\"This\"",
            getArtifactsDir() + "XpsSaveOptions.OptimizeOutput.xps", "1.fpage");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "optimizeOutputDataProvider")
	public static Object[][] optimizeOutputDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}
}
