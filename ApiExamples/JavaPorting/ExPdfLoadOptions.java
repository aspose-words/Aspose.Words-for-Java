package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.PdfLoadOptions;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import org.testng.Assert;
import org.testng.annotations.DataProvider;


@Test
class ExPdfLoadOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test (dataProvider = "skipPdfImagesDataProvider")
    public void skipPdfImages(boolean isSkipPdfImages) throws Exception
    {
        //ExStart
        //ExFor:PdfLoadOptions.SkipPdfImages
        //ExSummary:Shows how to skip images during loading PDF files.
        PdfLoadOptions options = new PdfLoadOptions();
        options.setSkipPdfImages(isSkipPdfImages);
        
        Document doc = new Document(getMyDir() + "Images.pdf", options);
        NodeCollection shapeCollection = doc.getChildNodes(NodeType.SHAPE, true);

        if (isSkipPdfImages)
        {
            Assert.assertEquals(shapeCollection.getCount(), 0);
        }
        else
        {
            Assert.assertNotEquals(shapeCollection.getCount(), 0);
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "skipPdfImagesDataProvider")
	public static Object[][] skipPdfImagesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}
}

