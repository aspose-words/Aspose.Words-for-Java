// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.TableContentAlignment;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.MarkdownSaveOptions;
import com.aspose.words.Document;
import com.aspose.words.Table;
import org.testng.Assert;
import com.aspose.ms.System.IO.Directory;
import com.aspose.words.IImageSavingCallback;
import com.aspose.words.ImageSavingArgs;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.IO.File;
import com.aspose.words.MarkdownListExportMode;
import com.aspose.ms.System.IO.Path;
import com.aspose.words.Underline;
import org.testng.annotations.DataProvider;


@Test
class ExMarkdownSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test (dataProvider = "markdownDocumentTableContentAlignmentDataProvider")
    public void markdownDocumentTableContentAlignment(/*TableContentAlignment*/int tableContentAlignment) throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        builder.insertCell();
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        builder.write("Cell1");
        builder.insertCell();
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.write("Cell2");

        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions(); { saveOptions.setTableContentAlignment(tableContentAlignment); }

        builder.getDocument().save(getArtifactsDir() + "MarkdownSaveOptions.MarkdownDocumentTableContentAlignment.md", saveOptions);

        Document doc = new Document(getArtifactsDir() + "MarkdownSaveOptions.MarkdownDocumentTableContentAlignment.md");
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        switch (tableContentAlignment)
        {
            case TableContentAlignment.AUTO:
                Assert.assertEquals(ParagraphAlignment.RIGHT,
                    table.getFirstRow().getCells().get(0).getFirstParagraph().getParagraphFormat().getAlignment());
                Assert.assertEquals(ParagraphAlignment.CENTER,
                    table.getFirstRow().getCells().get(1).getFirstParagraph().getParagraphFormat().getAlignment());
                break;
            case TableContentAlignment.LEFT:
                Assert.assertEquals(ParagraphAlignment.LEFT,
                    table.getFirstRow().getCells().get(0).getFirstParagraph().getParagraphFormat().getAlignment());
                Assert.assertEquals(ParagraphAlignment.LEFT,
                    table.getFirstRow().getCells().get(1).getFirstParagraph().getParagraphFormat().getAlignment());
                break;
            case TableContentAlignment.CENTER:
                Assert.assertEquals(ParagraphAlignment.CENTER,
                    table.getFirstRow().getCells().get(0).getFirstParagraph().getParagraphFormat().getAlignment());
                Assert.assertEquals(ParagraphAlignment.CENTER,
                    table.getFirstRow().getCells().get(1).getFirstParagraph().getParagraphFormat().getAlignment());
                break;
            case TableContentAlignment.RIGHT:
                Assert.assertEquals(ParagraphAlignment.RIGHT,
                    table.getFirstRow().getCells().get(0).getFirstParagraph().getParagraphFormat().getAlignment());
                Assert.assertEquals(ParagraphAlignment.RIGHT,
                    table.getFirstRow().getCells().get(1).getFirstParagraph().getParagraphFormat().getAlignment());
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "markdownDocumentTableContentAlignmentDataProvider")
	public static Object[][] markdownDocumentTableContentAlignmentDataProvider() throws Exception
	{
		return new Object[][]
		{
			{TableContentAlignment.LEFT},
			{TableContentAlignment.RIGHT},
			{TableContentAlignment.CENTER},
			{TableContentAlignment.AUTO},
		};
	}

    //ExStart
    //ExFor:MarkdownSaveOptions.ImageSavingCallback
    //ExFor:IImageSavingCallback
    //ExSummary:Shows how to rename the image name during saving into Markdown document.
    @Test //ExSkip
    public void renameImages() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // If we convert a document that contains images into Markdown, we will end up with one Markdown file which links to several images.
        // Each image will be in the form of a file in the local file system.
        // There is also a callback that can customize the name and file system location of each image.
        saveOptions.setImageSavingCallback(new SavedImageRename("MarkdownSaveOptions.HandleDocument.md"));

        // The ImageSaving() method of our callback will be run at this time.
        doc.save(getArtifactsDir() + "MarkdownSaveOptions.HandleDocument.md", saveOptions);

        Assert.AreEqual(1,
            Directory.getFiles(getArtifactsDir())
                .Where(s => s.StartsWith(ArtifactsDir + "MarkdownSaveOptions.HandleDocument.md shape"))
                .Count(f => f.EndsWith(".jpeg")));
        Assert.AreEqual(8,
            Directory.getFiles(getArtifactsDir())
                .Where(s => s.StartsWith(ArtifactsDir + "MarkdownSaveOptions.HandleDocument.md shape"))
                .Count(f => f.EndsWith(".png")));
    }

    /// <summary>
    /// Renames saved images that are produced when an Markdown document is saved.
    /// </summary>
    public static class SavedImageRename implements IImageSavingCallback
    {
        public SavedImageRename(String outFileName)
        {
            mOutFileName = outFileName;
        }

        public void /*IImageSavingCallback.*/imageSaving(ImageSavingArgs args) throws Exception
        {
            String imageFileName = $"{mOutFileName} shape {++mCount}, of type {args.CurrentShape.ShapeType}{Path.GetExtension(args.ImageFileName)}";

            args.setImageFileName(imageFileName);
            args.ImageStream = new FileStream(getArtifactsDir() + imageFileName, FileMode.CREATE);

            Assert.True(args.ImageStream.CanWrite);
            Assert.assertTrue(args.isImageAvailable());
            Assert.assertFalse(args.getKeepImageStreamOpen());
        }

        private int mCount;
        private /*final*/ String mOutFileName;
    }
    //ExEnd

    @Test (dataProvider = "exportImagesAsBase64DataProvider")
    public void exportImagesAsBase64(boolean exportImagesAsBase64) throws Exception
    {
        //ExStart
        //ExFor:MarkdownSaveOptions.ExportImagesAsBase64
        //ExSummary:Shows how to save a .md document with images embedded inside it.
        Document doc = new Document(getMyDir() + "Images.docx");

        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions(); { saveOptions.setExportImagesAsBase64(exportImagesAsBase64); }

        doc.save(getArtifactsDir() + "MarkdownSaveOptions.ExportImagesAsBase64.md", saveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "MarkdownSaveOptions.ExportImagesAsBase64.md");

        Assert.assertTrue(exportImagesAsBase64
            ? outDocContents.contains("data:image/jpeg;base64")
            : outDocContents.contains("MarkdownSaveOptions.ExportImagesAsBase64.001.jpeg"));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportImagesAsBase64DataProvider")
	public static Object[][] exportImagesAsBase64DataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "listExportModeDataProvider")
    public void listExportMode(/*MarkdownListExportMode*/int markdownListExportMode) throws Exception
    {
        //ExStart
        //ExFor:MarkdownSaveOptions.ListExportMode
        //ExSummary:Shows how to list items will be written to the markdown document.
        Document doc = new Document(getMyDir() + "List item.docx");

        // Use MarkdownListExportMode.PlainText or MarkdownListExportMode.MarkdownSyntax to export list.
        MarkdownSaveOptions options = new MarkdownSaveOptions(); { options.setListExportMode(markdownListExportMode); }
        doc.save(getArtifactsDir() + "MarkdownSaveOptions.ListExportMode.md", options);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "listExportModeDataProvider")
	public static Object[][] listExportModeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{MarkdownListExportMode.PLAIN_TEXT},
			{MarkdownListExportMode.MARKDOWN_SYNTAX},
		};
	}

    @Test
    public void imagesFolder() throws Exception
    {
        //ExStart
        //ExFor:MarkdownSaveOptions.ImagesFolder
        //ExFor:MarkdownSaveOptions.ImagesFolderAlias
        //ExSummary:Shows how to specifies the name of the folder used to construct image URIs.
        DocumentBuilder builder = new DocumentBuilder();

        builder.writeln("Some image below:");
        builder.insertImage(getImageDir() + "Logo.jpg");

        String imagesFolder = Path.combine(getArtifactsDir(), "ImagesDir");
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        // Use the "ImagesFolder" property to assign a folder in the local file system into which
        // Aspose.Words will save all the document's linked images.
        saveOptions.setImagesFolder(imagesFolder);
        // Use the "ImagesFolderAlias" property to use this folder
        // when constructing image URIs instead of the images folder's name.
        saveOptions.setImagesFolderAlias("http://example.com/images");

        builder.getDocument().save(getArtifactsDir() + "MarkdownSaveOptions.ImagesFolder.md", saveOptions);
        //ExEnd

        String[] dirFiles = Directory.getFiles(imagesFolder, "MarkdownSaveOptions.ImagesFolder.001.jpeg");
        Assert.assertEquals(1, dirFiles.length);
        Document doc = new Document(getArtifactsDir() + "MarkdownSaveOptions.ImagesFolder.md");
        doc.getText().contains("http://example.com/images/MarkdownSaveOptions.ImagesFolder.001.jpeg");
    }

    @Test
    public void exportUnderlineFormatting() throws Exception
    {
        //ExStart:ExportUnderlineFormatting
        //GistId:eeeec1fbf118e95e7df3f346c91ed726
        //ExFor:MarkdownSaveOptions.ExportUnderlineFormatting
        //ExSummary:Shows how to export underline formatting as ++.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.setUnderline(Underline.SINGLE);
        builder.write("Lorem ipsum. Dolor sit amet.");

        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions(); { saveOptions.setExportUnderlineFormatting(true); }
        doc.save(getArtifactsDir() + "MarkdownSaveOptions.ExportUnderlineFormatting.md", saveOptions);
        //ExEnd:ExportUnderlineFormatting
    }
}

