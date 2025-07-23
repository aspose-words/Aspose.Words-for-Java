package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

@Test
class ExPdfLoadOptions extends ApiExampleBase {
    @Test (dataProvider = "skipPdfImagesDataProvider")
    public void skipPdfImages(boolean isSkipPdfImages) throws Exception {
        //ExStart
        //ExFor:PdfLoadOptions
        //ExFor:PdfLoadOptions.SkipPdfImages
        //ExFor:PdfLoadOptions.PageIndex
        //ExFor:PdfLoadOptions.PageCount
        //ExSummary:Shows how to skip images during loading PDF files.
        PdfLoadOptions options = new PdfLoadOptions();
        options.setSkipPdfImages(isSkipPdfImages);
        options.setPageIndex(0);
        options.setPageCount(1);

        Document doc = new Document(getMyDir() + "Images.pdf", options);
        NodeCollection shapeCollection = doc.getChildNodes(NodeType.SHAPE, true);

        if (isSkipPdfImages)
            Assert.assertEquals(shapeCollection.getCount(), 0);
        else
            Assert.assertNotEquals(shapeCollection.getCount(), 0);
        //ExEnd
    }

	@DataProvider(name = "skipPdfImagesDataProvider")
	public static Object[][] skipPdfImagesDataProvider() {
		return new Object[][]
		{
			{true},
			{false},
		};
	}
}

