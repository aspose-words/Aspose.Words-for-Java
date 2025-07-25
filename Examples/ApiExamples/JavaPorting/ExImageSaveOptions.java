// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.PageSet;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.ms.System.IO.File;
import com.aspose.words.GraphicsQualityOptions;
import com.aspose.ms.System.Drawing.Drawing2D.SmoothingMode;
import com.aspose.ms.System.Drawing.Text.TextRenderingHint;
import com.aspose.ms.System.Drawing.Drawing2D.CompositingMode;
import com.aspose.ms.System.Drawing.Drawing2D.InterpolationMode;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.words.ShapeRenderer;
import com.aspose.words.MetafileRenderingMode;
import com.aspose.ms.System.Drawing.msSize;
import java.util.ArrayList;
import com.aspose.ms.System.IO.Directory;
import com.aspose.words.ImageColorMode;
import com.aspose.ms.System.IO.FileInfo;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.ImagePixelFormat;
import com.aspose.words.TiffCompression;
import com.aspose.words.ImageBinarizationMethod;
import com.aspose.words.PageRange;
import com.aspose.words.ImlRenderingMode;
import com.aspose.words.MultiPageLayout;
import org.testng.annotations.DataProvider;


@Test
class ExImageSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void onePage() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:FixedPageSaveOptions
        //ExFor:ImageSaveOptions.PageSet
        //ExFor:PageSet
        //ExFor:PageSet.#ctor(Int32)
        //ExSummary:Shows how to render one page from a document to a JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2.");
        builder.insertImage(getImageDir() + "Logo.jpg");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 3.");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.JPEG);
        // Set the "PageSet" to "1" to select the second page via
        // the zero-based index to start rendering the document from.
        options.setPageSet(new PageSet(1));

        // When we save the document to the JPEG format, Aspose.Words only renders one page.
        // This image will contain one page starting from page two,
        // which will just be the second page of the original document.
        doc.save(getArtifactsDir() + "ImageSaveOptions.OnePage.jpg", options);
        //ExEnd

        TestUtil.verifyImage(816, 1056, getArtifactsDir() + "ImageSaveOptions.OnePage.jpg");
    }

    @Test (dataProvider = "rendererDataProvider")
    public void renderer(boolean useGdiEmfRenderer) throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions.UseGdiEmfRenderer
        //ExSummary:Shows how to choose a renderer when converting a document to .emf.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("Hello world!");
        builder.insertImage(getImageDir() + "Logo.jpg");

        // When we save the document as an EMF image, we can pass a SaveOptions object to select a renderer for the image.
        // If we set the "UseGdiEmfRenderer" flag to "true", Aspose.Words will use the GDI+ renderer.
        // If we set the "UseGdiEmfRenderer" flag to "false", Aspose.Words will use its own metafile renderer.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.EMF);
        saveOptions.setUseGdiEmfRenderer(useGdiEmfRenderer);

        doc.save(getArtifactsDir() + "ImageSaveOptions.Renderer.emf", saveOptions);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "rendererDataProvider")
	public static Object[][] rendererDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void pageSet() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions.PageSet
        //ExSummary:Shows how to specify which page in a document to render as an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("Hello world! This is page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("This is page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("This is page 3.");

        Assert.assertEquals(3, doc.getPageCount());

        // When we save the document as an image, Aspose.Words only renders the first page by default.
        // We can pass a SaveOptions object to specify a different page to render.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.GIF);
        // Render every page of the document to a separate image file.
        for (int i = 1; i <= doc.getPageCount(); i++)
        {
            saveOptions.setPageSet(new PageSet(1));

            doc.save(getArtifactsDir() + $"ImageSaveOptions.PageIndex.Page {i}.gif", saveOptions);
        }
        //ExEnd

        TestUtil.verifyImage(816, 1056, getArtifactsDir() + "ImageSaveOptions.PageIndex.Page 1.gif");
        TestUtil.verifyImage(816, 1056, getArtifactsDir() + "ImageSaveOptions.PageIndex.Page 2.gif");
        TestUtil.verifyImage(816, 1056, getArtifactsDir() + "ImageSaveOptions.PageIndex.Page 3.gif");
        Assert.assertFalse(File.exists(getArtifactsDir() + "ImageSaveOptions.PageIndex.Page 4.gif"));
    }

    @Test
    public void graphicsQuality() throws Exception
    {
        //ExStart
        //ExFor:GraphicsQualityOptions
        //ExFor:GraphicsQualityOptions.CompositingMode
        //ExFor:GraphicsQualityOptions.CompositingQuality
        //ExFor:GraphicsQualityOptions.InterpolationMode
        //ExFor:GraphicsQualityOptions.StringFormat
        //ExFor:GraphicsQualityOptions.SmoothingMode
        //ExFor:GraphicsQualityOptions.TextRenderingHint
        //ExFor:ImageSaveOptions.GraphicsQualityOptions
        //ExSummary:Shows how to set render quality options while converting documents to image formats. 
        Document doc = new Document(getMyDir() + "Rendering.docx");

        GraphicsQualityOptions qualityOptions = new GraphicsQualityOptions();
        {
            qualityOptions.setSmoothingMode(SmoothingMode.ANTI_ALIAS);
            qualityOptions.setTextRenderingHint(TextRenderingHint.CLEAR_TYPE_GRID_FIT);
            qualityOptions.setCompositingMode(CompositingMode.SOURCE_OVER);
            qualityOptions.setCompositingQuality(CompositingQuality.HighQuality);
            qualityOptions.setInterpolationMode(InterpolationMode.HIGH);
            qualityOptions.setStringFormat(StringFormat.GenericTypographic);
        }

        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.JPEG);
        saveOptions.setGraphicsQualityOptions(qualityOptions);

        doc.save(getArtifactsDir() + "ImageSaveOptions.GraphicsQuality.jpg", saveOptions);
        //ExEnd

        TestUtil.verifyImage(794, 1122, getArtifactsDir() + "ImageSaveOptions.GraphicsQuality.jpg");
    }

    @Test
    public void useTileFlipMode() throws Exception
    {
        //ExStart
        //ExFor:GraphicsQualityOptions.UseTileFlipMode
        //ExSummary:Shows how to prevent the white line appears when rendering with a high resolution.
        Document doc = new Document(getMyDir() + "Shape high dpi.docx");

        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        ShapeRenderer renderer = shape.getShapeRenderer();

        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.PNG);
        {
            saveOptions.setResolution(500f); saveOptions.setGraphicsQualityOptions(new GraphicsQualityOptions()); { saveOptions.getGraphicsQualityOptions().setUseTileFlipMode(true); }
        }
        renderer.save(getArtifactsDir() + "ImageSaveOptions.UseTileFlipMode.png", saveOptions);
        //ExEnd
    }

    @Test (groups = "SkipMono", dataProvider = "windowsMetaFileDataProvider")
    public void windowsMetaFile(/*MetafileRenderingMode*/int metafileRenderingMode) throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions.MetafileRenderingOptions
        //ExFor:MetafileRenderingOptions.UseGdiRasterOperationsEmulation
        //ExSummary:Shows how to set the rendering mode when saving documents with Windows Metafile images to other image formats. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImageDir() + "Windows MetaFile.wmf");

        // When we save the document as an image, we can pass a SaveOptions object to
        // determine how the saving operation will process Windows Metafiles in the document.
        // If we set the "RenderingMode" property to "MetafileRenderingMode.Vector",
        // or "MetafileRenderingMode.VectorWithFallback", we will render all metafiles as vector graphics.
        // If we set the "RenderingMode" property to "MetafileRenderingMode.Bitmap", we will render all metafiles as bitmaps.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
        options.getMetafileRenderingOptions().setRenderingMode(metafileRenderingMode);
        // Aspose.Words uses GDI+ for raster operations emulation, when value is set to true.
        options.getMetafileRenderingOptions().setUseGdiRasterOperationsEmulation(true);

        doc.save(getArtifactsDir() + "ImageSaveOptions.WindowsMetaFile.png", options);
        //ExEnd

        TestUtil.verifyImage(816, 1056, getArtifactsDir() + "ImageSaveOptions.WindowsMetaFile.png");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "windowsMetaFileDataProvider")
	public static Object[][] windowsMetaFileDataProvider() throws Exception
	{
		return new Object[][]
		{
			{MetafileRenderingMode.VECTOR},
			{MetafileRenderingMode.BITMAP},
			{MetafileRenderingMode.VECTOR_WITH_FALLBACK},
		};
	}

    @Test (groups = "SkipMono")
    public void pageByPage() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:FixedPageSaveOptions
        //ExFor:ImageSaveOptions.PageSet
        //ExFor:ImageSaveOptions.ImageSize
        //ExSummary:Shows how to render every page of a document to a separate TIFF image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2.");
        builder.insertImage(getImageDir() + "Logo.jpg");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 3.");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.TIFF);

        for (int i = 0; i < doc.getPageCount(); i++)
        {
            // Set the "PageSet" property to the number of the first page from
            // which to start rendering the document from.
            options.setPageSet(new PageSet(i));
            // Export page at 2325x5325 pixels and 600 dpi.
            options.setResolution(600f);
            options.setImageSizeInternal(msSize.ctor(2325, 5325));

            doc.save(getArtifactsDir() + $"ImageSaveOptions.PageByPage.{i + 1}.tiff", options);
        }
        //ExEnd

        ArrayList<String> imageFileNames = Directory.getFiles(getArtifactsDir(), "*.tiff")
            .Where(item => item.Contains("ImageSaveOptions.PageByPage.") && item.EndsWith(".tiff")).ToList();
        Assert.assertEquals(3, imageFileNames.size());
    }

    @Test (dataProvider = "colorModeDataProvider")
    public void colorMode(/*ImageColorMode*/int imageColorMode) throws Exception
    {
        //ExStart
        //ExFor:ImageColorMode
        //ExFor:ImageSaveOptions.ImageColorMode
        //ExSummary:Shows how to set a color mode when rendering documents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("Hello world!");
        builder.insertImage(getImageDir() + "Logo.jpg");

        // When we save the document as an image, we can pass a SaveOptions object to
        // select a color mode for the image that the saving operation will generate.
        // If we set the "ImageColorMode" property to "ImageColorMode.BlackAndWhite",
        // the saving operation will apply grayscale color reduction while rendering the document.
        // If we set the "ImageColorMode" property to "ImageColorMode.Grayscale", 
        // the saving operation will render the document into a monochrome image.
        // If we set the "ImageColorMode" property to "None", the saving operation will apply the default method
        // and preserve all the document's colors in the output image.
        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setImageColorMode(imageColorMode);

        doc.save(getArtifactsDir() + "ImageSaveOptions.ColorMode.png", imageSaveOptions);
        //ExEnd

        long testedImageLength = new FileInfo(getArtifactsDir() + "ImageSaveOptions.ColorMode.png").getLength();

        switch (imageColorMode)
        {
            case ImageColorMode.NONE:
                Assert.assertTrue(testedImageLength < 175000);
                break;
            case ImageColorMode.GRAYSCALE:
                Assert.assertTrue(testedImageLength < 90000);
                break;
            case ImageColorMode.BLACK_AND_WHITE:
                Assert.assertTrue(testedImageLength < 15000);
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "colorModeDataProvider")
	public static Object[][] colorModeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{ImageColorMode.BLACK_AND_WHITE},
			{ImageColorMode.GRAYSCALE},
			{ImageColorMode.NONE},
		};
	}

    @Test
    public void paperColor() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.PaperColor
        //ExSummary:Renders a page of a Word document into an image with transparent or colored background.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Times New Roman");
        builder.getFont().setSize(24.0);
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        builder.insertImage(getImageDir() + "Logo.jpg");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.PNG);
        // Set the "PaperColor" property to a transparent color to apply a transparent
        // background to the document while rendering it to an image.
        imgOptions.setPaperColor(msColor.getTransparent());

        doc.save(getArtifactsDir() + "ImageSaveOptions.PaperColor.Transparent.png", imgOptions);

        // Set the "PaperColor" property to an opaque color to apply that color
        // as the background of the document as we render it to an image.
        imgOptions.setPaperColor(msColor.getLightCoral());

        doc.save(getArtifactsDir() + "ImageSaveOptions.PaperColor.LightCoral.png", imgOptions);
        //ExEnd

        TestUtil.imageContainsTransparency(getArtifactsDir() + "ImageSaveOptions.PaperColor.Transparent.png");
        Assert.<AssertionError>Throws(() =>
            TestUtil.imageContainsTransparency(getArtifactsDir() + "ImageSaveOptions.PaperColor.LightCoral.png"));
    }

    @Test (dataProvider = "pixelFormatDataProvider")
    public void pixelFormat(/*ImagePixelFormat*/int imagePixelFormat) throws Exception
    {
        //ExStart
        //ExFor:ImagePixelFormat
        //ExFor:ImageSaveOptions.Clone
        //ExFor:ImageSaveOptions.PixelFormat
        //ExSummary:Shows how to select a bit-per-pixel rate with which to render a document to an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("Hello world!");
        builder.insertImage(getImageDir() + "Logo.jpg");

        // When we save the document as an image, we can pass a SaveOptions object to
        // select a pixel format for the image that the saving operation will generate.
        // Various bit per pixel rates will affect the quality and file size of the generated image.
        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPixelFormat(imagePixelFormat);

        // We can clone ImageSaveOptions instances.
        Assert.Is.Not.EqualTo(imageSaveOptions)imageSaveOptions.deepClone());

        doc.save(getArtifactsDir() + "ImageSaveOptions.PixelFormat.png", imageSaveOptions);
        //ExEnd

        long testedImageLength = new FileInfo(getArtifactsDir() + "ImageSaveOptions.PixelFormat.png").getLength();

        switch (imagePixelFormat)
        {
            case ImagePixelFormat.FORMAT_1_BPP_INDEXED:
                Assert.assertTrue(testedImageLength < 2500);
                break;
            case ImagePixelFormat.FORMAT_16_BPP_RGB_565:
                Assert.assertTrue(testedImageLength < 104000);
                break;
            case ImagePixelFormat.FORMAT_16_BPP_RGB_555:
                Assert.assertTrue(testedImageLength < 88000);
                break;
            case ImagePixelFormat.FORMAT_24_BPP_RGB:
                Assert.assertTrue(testedImageLength < 160000);
                break;
            case ImagePixelFormat.FORMAT_32_BPP_RGB:
            case ImagePixelFormat.FORMAT_32_BPP_ARGB:
                Assert.assertTrue(testedImageLength < 175000);
                break;
            case ImagePixelFormat.FORMAT_48_BPP_RGB:
                Assert.assertTrue(testedImageLength < 212000);
                break;
            case ImagePixelFormat.FORMAT_64_BPP_ARGB:
            case ImagePixelFormat.FORMAT_64_BPP_P_ARGB:
                Assert.assertTrue(testedImageLength < 239000);
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "pixelFormatDataProvider")
	public static Object[][] pixelFormatDataProvider() throws Exception
	{
		return new Object[][]
		{
			{ImagePixelFormat.FORMAT_1_BPP_INDEXED},
			{ImagePixelFormat.FORMAT_16_BPP_RGB_555},
			{ImagePixelFormat.FORMAT_16_BPP_RGB_565},
			{ImagePixelFormat.FORMAT_24_BPP_RGB},
			{ImagePixelFormat.FORMAT_32_BPP_RGB},
			{ImagePixelFormat.FORMAT_32_BPP_ARGB},
			{ImagePixelFormat.FORMAT_32_BPP_P_ARGB},
			{ImagePixelFormat.FORMAT_48_BPP_RGB},
			{ImagePixelFormat.FORMAT_64_BPP_ARGB},
			{ImagePixelFormat.FORMAT_64_BPP_P_ARGB},
		};
	}

    @Test (groups = "SkipMono")
    public void floydSteinbergDithering() throws Exception
    {
        //ExStart
        //ExFor:ImageBinarizationMethod
        //ExFor:ImageSaveOptions.ThresholdForFloydSteinbergDithering
        //ExFor:ImageSaveOptions.TiffBinarizationMethod
        //ExSummary:Shows how to set the TIFF binarization error threshold when using the Floyd-Steinberg method to render a TIFF image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("Hello world!");
        builder.insertImage(getImageDir() + "Logo.jpg");

        // When we save the document as a TIFF, we can pass a SaveOptions object to
        // adjust the dithering that Aspose.Words will apply when rendering this image.
        // The default value of the "ThresholdForFloydSteinbergDithering" property is 128.
        // Higher values tend to produce darker images.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.TIFF);
        {
            options.setTiffCompression(TiffCompression.CCITT_3);
            options.setTiffBinarizationMethod(ImageBinarizationMethod.FLOYD_STEINBERG_DITHERING);
            options.setThresholdForFloydSteinbergDithering((byte) 240);
        }

        doc.save(getArtifactsDir() + "ImageSaveOptions.FloydSteinbergDithering.tiff", options);
        //ExEnd

        ArrayList<String> imageFileNames = Directory.getFiles(getArtifactsDir(), "*.tiff")
            .Where(item => item.Contains("ImageSaveOptions.FloydSteinbergDithering.") && item.EndsWith(".tiff")).ToList();
        Assert.assertEquals(1, imageFileNames.size());
    }

    @Test
    public void editImage() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions.HorizontalResolution
        //ExFor:ImageSaveOptions.ImageBrightness
        //ExFor:ImageSaveOptions.ImageContrast
        //ExFor:ImageSaveOptions.SaveFormat
        //ExFor:ImageSaveOptions.Scale
        //ExFor:ImageSaveOptions.VerticalResolution
        //ExSummary:Shows how to edit the image while Aspose.Words converts a document to one.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("Hello world!");
        builder.insertImage(getImageDir() + "Logo.jpg");

        // When we save the document as an image, we can pass a SaveOptions object to
        // edit the image while the saving operation renders it.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
        {
            // We can adjust these properties to change the image's brightness and contrast.
            // Both are on a 0-1 scale and are at 0.5 by default.
            options.setImageBrightness(0.3f);
            options.setImageContrast(0.7f);

            // We can adjust horizontal and vertical resolution with these properties.
            // This will affect the dimensions of the image.
            // The default value for these properties is 96.0, for a resolution of 96dpi.
            options.setHorizontalResolution(72f);
            options.setVerticalResolution(72f);

            // We can scale the image using this property. The default value is 1.0, for scaling of 100%.
            // We can use this property to negate any changes in image dimensions that changing the resolution would cause.
            options.setScale(96f / 72f);
        }

        doc.save(getArtifactsDir() + "ImageSaveOptions.EditImage.png", options);
        //ExEnd

        TestUtil.verifyImage(817, 1057, getArtifactsDir() + "ImageSaveOptions.EditImage.png");
    }

    @Test
    public void jpegQuality() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:FixedPageSaveOptions.JpegQuality
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.#ctor
        //ExFor:ImageSaveOptions.JpegQuality
        //ExSummary:Shows how to configure compression while saving a document as a JPEG.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertImage(getImageDir() + "Logo.jpg");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);
        // Set the "JpegQuality" property to "10" to use stronger compression when rendering the document.
        // This will reduce the file size of the document, but the image will display more prominent compression artifacts.
        imageOptions.setJpegQuality(10);
        doc.save(getArtifactsDir() + "ImageSaveOptions.JpegQuality.HighCompression.jpg", imageOptions);

        // Set the "JpegQuality" property to "100" to use weaker compression when rending the document.
        // This will improve the quality of the image at the cost of an increased file size.
        imageOptions.setJpegQuality(100);
        doc.save(getArtifactsDir() + "ImageSaveOptions.JpegQuality.HighQuality.jpg", imageOptions);
        //ExEnd

        Assert.assertTrue(new FileInfo(getArtifactsDir() + "ImageSaveOptions.JpegQuality.HighCompression.jpg").getLength() < 18000);
        Assert.assertTrue(new FileInfo(getArtifactsDir() + "ImageSaveOptions.JpegQuality.HighQuality.jpg").getLength() < 75000);
    }

    @Test (groups = "SkipMono", dataProvider = "tiffImageCompressionDataProvider")
    public void tiffImageCompression(/*TiffCompression*/int tiffCompression) throws Exception
    {
        //ExStart
        //ExFor:TiffCompression
        //ExFor:ImageSaveOptions.TiffCompression
        //ExSummary:Shows how to select the compression scheme to apply to a document that we convert into a TIFF image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImageDir() + "Logo.jpg");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.TIFF);
        // Set the "TiffCompression" property to "TiffCompression.None" to apply no compression while saving,
        // which may result in a very large output file.
        // Set the "TiffCompression" property to "TiffCompression.Rle" to apply RLE compression
        // Set the "TiffCompression" property to "TiffCompression.Lzw" to apply LZW compression.
        // Set the "TiffCompression" property to "TiffCompression.Ccitt3" to apply CCITT3 compression.
        // Set the "TiffCompression" property to "TiffCompression.Ccitt4" to apply CCITT4 compression.
        options.setTiffCompression(tiffCompression);

        doc.save(getArtifactsDir() + "ImageSaveOptions.TiffImageCompression.tiff", options);
        //ExEnd

        long testedImageLength = new FileInfo(getArtifactsDir() + "ImageSaveOptions.TiffImageCompression.tiff").getLength();

        switch (tiffCompression)
        {
            case TiffCompression.NONE:
                Assert.assertTrue(testedImageLength < 3450000);
                break;
            case TiffCompression.RLE:
                Assert.assertTrue(testedImageLength < 687000);
                break;
            case TiffCompression.LZW:
                Assert.assertTrue(testedImageLength < 250000);
                break;
            case TiffCompression.CCITT_3:
                Assert.assertTrue(testedImageLength < 8300);
                break;
            case TiffCompression.CCITT_4:
                Assert.assertTrue(testedImageLength < 1700);
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "tiffImageCompressionDataProvider")
	public static Object[][] tiffImageCompressionDataProvider() throws Exception
	{
		return new Object[][]
		{
			{TiffCompression.NONE},
			{TiffCompression.RLE},
			{TiffCompression.LZW},
			{TiffCompression.CCITT_3},
			{TiffCompression.CCITT_4},
		};
	}

    @Test
    public void resolution() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.Resolution
        //ExSummary:Shows how to specify a resolution while rendering a document to PNG.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Times New Roman");
        builder.getFont().setSize(24.0);
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        builder.insertImage(getImageDir() + "Logo.jpg");

        // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method renders the document into an image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);

        // Set the "Resolution" property to "72" to render the document in 72dpi.
        options.setResolution(72f);
        doc.save(getArtifactsDir() + "ImageSaveOptions.Resolution.72dpi.png", options);

        // Set the "Resolution" property to "300" to render the document in 300dpi.
        options.setResolution(300f);
        doc.save(getArtifactsDir() + "ImageSaveOptions.Resolution.300dpi.png", options);
        //ExEnd

        TestUtil.verifyImage(612, 792, getArtifactsDir() + "ImageSaveOptions.Resolution.72dpi.png");
        TestUtil.verifyImage(2550, 3300, getArtifactsDir() + "ImageSaveOptions.Resolution.300dpi.png");
    }

    @Test
    public void exportVariousPageRanges() throws Exception
    {
        //ExStart
        //ExFor:PageSet.#ctor(PageRange[])
        //ExFor:PageRange
        //ExFor:PageRange.#ctor(int, int)
        //ExFor:ImageSaveOptions.PageSet
        //ExSummary:Shows how to extract pages based on exact page ranges.
        Document doc = new Document(getMyDir() + "Images.docx");

        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.TIFF);
        PageSet pageSet = new PageSet(new PageRange(1, 1), new PageRange(2, 3), new PageRange(1, 3),
            new PageRange(2, 4), new PageRange(1, 1));

        imageOptions.setPageSet(pageSet);
        doc.save(getArtifactsDir() + "ImageSaveOptions.ExportVariousPageRanges.tiff", imageOptions);
        //ExEnd
    }

    @Test
    public void renderInkObject() throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.ImlRenderingMode
        //ExFor:ImlRenderingMode
        //ExSummary:Shows how to render Ink object.
        Document doc = new Document(getMyDir() + "Ink object.docx");

        // Set 'ImlRenderingMode.InkML' ignores fall-back shape of ink (InkML) object and renders InkML itself.
        // If the rendering result is unsatisfactory,
        // please use 'ImlRenderingMode.Fallback' to get a result similar to previous versions.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.JPEG);
        {
            saveOptions.setImlRenderingMode(ImlRenderingMode.INK_ML);
        }

        doc.save(getArtifactsDir() + "ImageSaveOptions.RenderInkObject.jpeg", saveOptions);
        //ExEnd
    }

    @Test
    public void gridLayout() throws Exception
    {
        //ExStart:GridLayout
        //GistId:70330eacdfc2e253f00a9adea8972975
        //ExFor:ImageSaveOptions.PageLayout
        //ExFor:MultiPageLayout
        //ExSummary:Shows how to save the document into JPG image with multi-page layout settings.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.JPEG);
        // Set up a grid layout with:
        // - 3 columns per row.
        // - 10pts spacing between pages (horizontal and vertical).
        options.setPageLayout(MultiPageLayout.grid(3, 10f, 10f));

        // Alternative layouts:
        // options.PageLayout = MultiPageLayout.Horizontal(10);
        // options.PageLayout = MultiPageLayout.Vertical(10);

        // Customize the background and border.
        options.getPageLayout().setBackColor(msColor.getLightGray());
        options.getPageLayout().setBorderColor(Color.BLUE);
        options.getPageLayout().setBorderWidth(2f);

        doc.save(getArtifactsDir() + "ImageSaveOptions.GridLayout.jpg", options);
        //ExEnd:GridLayout
    }
}

