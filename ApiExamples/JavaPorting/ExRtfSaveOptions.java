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
import com.aspose.words.RtfSaveOptions;
import org.testng.Assert;
import com.aspose.words.SaveFormat;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;
import com.aspose.words.ImageType;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import org.testng.annotations.DataProvider;


@Test
public class ExRtfSaveOptions extends ApiExampleBase
{
    @Test (dataProvider = "exportImagesDataProvider")
    public void exportImages(boolean exportImagesForOldReaders) throws Exception
    {
        //ExStart
        //ExFor:RtfSaveOptions
        //ExFor:RtfSaveOptions.ExportCompactSize
        //ExFor:RtfSaveOptions.ExportImagesForOldReaders
        //ExFor:RtfSaveOptions.SaveFormat
        //ExSummary:Shows how to save a document to .rtf with custom options.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
        RtfSaveOptions options = new RtfSaveOptions();

        Assert.assertEquals(SaveFormat.RTF, options.getSaveFormat());

        // Set the "ExportCompactSize" property to "true" to
        // reduce the saved document's size at the cost of right-to-left text compatibility.
        options.setExportCompactSize(true);

        // Set the "ExportImagesFotOldReaders" property to "true" to use extra keywords to ensure that our document is
        // compatible with pre-Microsoft Word 97 readers and WordPad.
        // Set the "ExportImagesFotOldReaders" property to "false" to reduce the size of the document,
        // but prevent old readers from being able to read any non-metafile or BMP images that the document may contain.
        options.setExportImagesForOldReaders(exportImagesForOldReaders);

        doc.save(getArtifactsDir() + "RtfSaveOptions.ExportImages.rtf", options);
        //ExEnd

        if (exportImagesForOldReaders)
        {
            TestUtil.fileContainsString("nonshppict", getArtifactsDir() + "RtfSaveOptions.ExportImages.rtf");
            TestUtil.fileContainsString("shprslt", getArtifactsDir() + "RtfSaveOptions.ExportImages.rtf");
        }
        else
        {
            if (!isRunningOnMono())
            {
                Assert.<AssertionError>Throws(() =>
                    TestUtil.fileContainsString("nonshppict", getArtifactsDir() + "RtfSaveOptions.ExportImages.rtf"));
                Assert.<AssertionError>Throws(() =>
                    TestUtil.fileContainsString("shprslt", getArtifactsDir() + "RtfSaveOptions.ExportImages.rtf"));
            }
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportImagesDataProvider")
	public static Object[][] exportImagesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (groups = "SkipMono", dataProvider = "saveImagesAsWmfDataProvider")
    public void saveImagesAsWmf(boolean saveImagesAsWmf) throws Exception
    {
        //ExStart
        //ExFor:RtfSaveOptions.SaveImagesAsWmf
        //ExSummary:Shows how to convert all images in a document to the Windows Metafile format as we save the document as an RTF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Jpeg image:");
        Shape imageShape = builder.insertImage(getImageDir() + "Logo.jpg");

        Assert.assertEquals(ImageType.JPEG, imageShape.getImageData().getImageType());

        builder.insertParagraph();
        builder.writeln("Png image:");
        imageShape = builder.insertImage(getImageDir() + "Transparent background logo.png");

        Assert.assertEquals(ImageType.PNG, imageShape.getImageData().getImageType());

        // Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
        RtfSaveOptions rtfSaveOptions = new RtfSaveOptions();

        // Set the "SaveImagesAsWmf" property to "true" to convert all images in the document to WMF as we save it to RTF.
        // Doing so will help readers such as WordPad to read our document.
        // Set the "SaveImagesAsWmf" property to "false" to preserve the original format of all images in the document
        // as we save it to RTF. This will preserve the quality of the images at the cost of compatibility with older RTF readers.
        rtfSaveOptions.setSaveImagesAsWmf(saveImagesAsWmf);

        doc.save(getArtifactsDir() + "RtfSaveOptions.SaveImagesAsWmf.rtf", rtfSaveOptions);

        doc = new Document(getArtifactsDir() + "RtfSaveOptions.SaveImagesAsWmf.rtf");

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        if (saveImagesAsWmf)
        {
            Assert.assertEquals(ImageType.WMF, ((Shape)shapes.get(0)).getImageData().getImageType());
            Assert.assertEquals(ImageType.WMF, ((Shape)shapes.get(1)).getImageData().getImageType());
        }
        else
        {
            Assert.assertEquals(ImageType.JPEG, ((Shape)shapes.get(0)).getImageData().getImageType());
            Assert.assertEquals(ImageType.PNG, ((Shape)shapes.get(1)).getImageData().getImageType());
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "saveImagesAsWmfDataProvider")
	public static Object[][] saveImagesAsWmfDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}
}
