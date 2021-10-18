// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.HtmlFixedSaveOptions;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.PsSaveOptions;
import com.aspose.words.SvgSaveOptions;
import com.aspose.words.XamlFixedSaveOptions;
import com.aspose.words.XpsSaveOptions;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.ms.System.IO.Directory;
import org.testng.Assert;
import com.aspose.words.IPageSavingCallback;
import com.aspose.words.PageSavingArgs;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.DocumentSplitCriteria;
import com.aspose.words.IDocumentPartSavingCallback;
import com.aspose.words.DocumentPartSavingArgs;
import com.aspose.words.IImageSavingCallback;
import com.aspose.words.ImageSavingArgs;
import com.aspose.words.CssStyleSheetType;
import com.aspose.words.ICssSavingCallback;
import com.aspose.words.CssSavingArgs;


@Test
class ExSavingCallback !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void checkThatAllMethodsArePresent()
    {
        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        htmlFixedSaveOptions.setPageSavingCallback(new CustomFileNamePageSavingCallback());

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSavingCallback(new CustomFileNamePageSavingCallback());

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setPageSavingCallback(new CustomFileNamePageSavingCallback());

        PsSaveOptions psSaveOptions = new PsSaveOptions();
        psSaveOptions.setPageSavingCallback(new CustomFileNamePageSavingCallback());

        SvgSaveOptions svgSaveOptions = new SvgSaveOptions();
        svgSaveOptions.setPageSavingCallback(new CustomFileNamePageSavingCallback());

        XamlFixedSaveOptions xamlFixedSaveOptions = new XamlFixedSaveOptions();
        xamlFixedSaveOptions.setPageSavingCallback(new CustomFileNamePageSavingCallback());

        XpsSaveOptions xpsSaveOptions = new XpsSaveOptions();
        xpsSaveOptions.setPageSavingCallback(new CustomFileNamePageSavingCallback());
    }

    //ExStart
    //ExFor:IPageSavingCallback
    //ExFor:IPageSavingCallback.PageSaving(PageSavingArgs)
    //ExFor:PageSavingArgs
    //ExFor:PageSavingArgs.PageFileName
    //ExFor:PageSavingArgs.KeepPageStreamOpen
    //ExFor:PageSavingArgs.PageIndex
    //ExFor:PageSavingArgs.PageStream
    //ExFor:FixedPageSaveOptions.PageSavingCallback
    //ExSummary:Shows how to use a callback to save a document to HTML page by page.
    @Test //ExSkip
    public void pageFileNames() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2.");
        builder.insertImage(getImageDir() + "Logo.jpg");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 3.");

        // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we convert the document to HTML.
        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();

        // We will save each page in this document to a separate HTML file in the local file system.
        // Set a callback that allows us to name each output HTML document.
        htmlFixedSaveOptions.setPageSavingCallback(new CustomFileNamePageSavingCallback());

        doc.save(getArtifactsDir() + "SavingCallback.PageFileNames.html", htmlFixedSaveOptions);

        String[] filePaths = Directory.getFiles(getArtifactsDir()).Where(
            s => s.StartsWith(ArtifactsDir + "SavingCallback.PageFileNames.Page_")).OrderBy(s => s).ToArray();

        Assert.assertEquals(3, filePaths.length);
    }

    /// <summary>
    /// Saves all pages to a file and directory specified within.
    /// </summary>
    private static class CustomFileNamePageSavingCallback implements IPageSavingCallback
    {
        public void pageSaving(PageSavingArgs args) throws Exception
        {
            String outFileName = $"{ArtifactsDir}SavingCallback.PageFileNames.Page_{args.PageIndex}.html";

            // Below are two ways of specifying where Aspose.Words will save each page of the document.
            // 1 -  Set a filename for the output page file:
            args.setPageFileName(outFileName);

            // 2 -  Create a custom stream for the output page file:
            args.setPageStreamInternal(new FileStream(outFileName, FileMode.CREATE));

            Assert.assertFalse(args.getKeepPageStreamOpen());
        }
    }
    //ExEnd

    //ExStart
    //ExFor:DocumentPartSavingArgs
    //ExFor:DocumentPartSavingArgs.Document
    //ExFor:DocumentPartSavingArgs.DocumentPartFileName
    //ExFor:DocumentPartSavingArgs.DocumentPartStream
    //ExFor:DocumentPartSavingArgs.KeepDocumentPartStreamOpen
    //ExFor:IDocumentPartSavingCallback
    //ExFor:IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs)
    //ExFor:IImageSavingCallback
    //ExFor:IImageSavingCallback.ImageSaving
    //ExFor:ImageSavingArgs
    //ExFor:ImageSavingArgs.ImageFileName
    //ExFor:HtmlSaveOptions
    //ExFor:HtmlSaveOptions.DocumentPartSavingCallback
    //ExFor:HtmlSaveOptions.ImageSavingCallback
    //ExSummary:Shows how to split a document into parts and save them.
    @Test //ExSkip
    public void documentPartsFileNames() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        String outFileName = "SavingCallback.DocumentPartsFileNames.html";

        // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we convert the document to HTML.
        HtmlSaveOptions options = new HtmlSaveOptions();

        // If we save the document normally, there will be one output HTML
        // document with all the source document's contents.
        // Set the "DocumentSplitCriteria" property to "DocumentSplitCriteria.SectionBreak" to
        // save our document to multiple HTML files: one for each section.
        options.setDocumentSplitCriteria(DocumentSplitCriteria.SECTION_BREAK);

        // Assign a custom callback to the "DocumentPartSavingCallback" property to alter the document part saving logic.
        options.setDocumentPartSavingCallback(new SavedDocumentPartRename(outFileName, options.getDocumentSplitCriteria()));

        // If we convert a document that contains images into html, we will end up with one html file which links to several images.
        // Each image will be in the form of a file in the local file system.
        // There is also a callback that can customize the name and file system location of each image.
        options.setImageSavingCallback(new SavedImageRename(outFileName));

        doc.save(getArtifactsDir() + outFileName, options);
    }

    /// <summary>
    /// Sets custom filenames for output documents that the saving operation splits a document into.
    /// </summary>
    private static class SavedDocumentPartRename implements IDocumentPartSavingCallback
    {
        public SavedDocumentPartRename(String outFileName, /*DocumentSplitCriteria*/int documentSplitCriteria)
        {
            mOutFileName = outFileName;
            mDocumentSplitCriteria = documentSplitCriteria;
        }

        public void /*IDocumentPartSavingCallback.*/documentPartSaving(DocumentPartSavingArgs args) throws Exception
        {
            // We can access the entire source document via the "Document" property.
            Assert.assertTrue(args.getDocument().getOriginalFileName().endsWith("Rendering.docx"));

            String partType = "";

            switch (mDocumentSplitCriteria)
            {
                case DocumentSplitCriteria.PAGE_BREAK:
                    partType = "Page";
                    break;
                case DocumentSplitCriteria.COLUMN_BREAK:
                    partType = "Column";
                    break;
                case DocumentSplitCriteria.SECTION_BREAK:
                    partType = "Section";
                    break;
                case DocumentSplitCriteria.HEADING_PARAGRAPH:
                    partType = "Paragraph from heading";
                    break;
            }

            String partFileName = $"{mOutFileName} part {++mCount}, of type {partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Below are two ways of specifying where Aspose.Words will save each part of the document.
            // 1 -  Set a filename for the output part file:
            args.setDocumentPartFileName(partFileName);

            // 2 -  Create a custom stream for the output part file:
            args.setDocumentPartStreamInternal(new FileStream(getArtifactsDir() + partFileName, FileMode.CREATE));

            Assert.assertTrue(args.getDocumentPartStreamInternal().canWrite());
            Assert.assertFalse(args.getKeepDocumentPartStreamOpen());
        }

        private int mCount;
        private /*final*/ String mOutFileName;
        private /*final*/ /*DocumentSplitCriteria*/int mDocumentSplitCriteria;
    }

    /// <summary>
    /// Sets custom filenames for image files that an HTML conversion creates.
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

            // Below are two ways of specifying where Aspose.Words will save each part of the document.
            // 1 -  Set a filename for the output image file:
            args.setImageFileName(imageFileName);

            // 2 -  Create a custom stream for the output image file:
            args.ImageStream = new FileStream(getArtifactsDir() + imageFileName, FileMode.CREATE);

            Assert.True(args.ImageStream.CanWrite);
            Assert.assertTrue(args.isImageAvailable());
            Assert.assertFalse(args.getKeepImageStreamOpen());
        }

        private int mCount;
        private /*final*/ String mOutFileName;
    }
    //ExEnd

    //ExStart
    //ExFor:CssSavingArgs
    //ExFor:CssSavingArgs.CssStream
    //ExFor:CssSavingArgs.Document
    //ExFor:CssSavingArgs.IsExportNeeded
    //ExFor:CssSavingArgs.KeepCssStreamOpen
    //ExFor:CssStyleSheetType
    //ExFor:HtmlSaveOptions.CssSavingCallback
    //ExFor:HtmlSaveOptions.CssStyleSheetFileName
    //ExFor:HtmlSaveOptions.CssStyleSheetType
    //ExFor:ICssSavingCallback
    //ExFor:ICssSavingCallback.CssSaving(CssSavingArgs)
    //ExSummary:Shows how to work with CSS stylesheets that an HTML conversion creates.
    @Test //ExSkip
    public void externalCssFilenames() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we convert the document to HTML.
        HtmlSaveOptions options = new HtmlSaveOptions();

        // Set the "CssStylesheetType" property to "CssStyleSheetType.External" to
        // accompany a saved HTML document with an external CSS stylesheet file.
        options.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);

        // Below are two ways of specifying directories and filenames for output CSS stylesheets.
        // 1 -  Use the "CssStyleSheetFileName" property to assign a filename to our stylesheet:
        options.setCssStyleSheetFileName(getArtifactsDir() + "SavingCallback.ExternalCssFilenames.css");

        // 2 -  Use a custom callback to name our stylesheet:
        options.setCssSavingCallback(new CustomCssSavingCallback(getArtifactsDir() + "SavingCallback.ExternalCssFilenames.css", true, false));

        doc.save(getArtifactsDir() + "SavingCallback.ExternalCssFilenames.html", options);
    }

    /// <summary>
    /// Sets a custom filename, along with other parameters for an external CSS stylesheet.
    /// </summary>
    private static class CustomCssSavingCallback implements ICssSavingCallback
    {
        public CustomCssSavingCallback(String cssDocFilename, boolean isExportNeeded, boolean keepCssStreamOpen)
        {
            mCssTextFileName = cssDocFilename;
            mIsExportNeeded = isExportNeeded;
            mKeepCssStreamOpen = keepCssStreamOpen;
        }

        public void cssSaving(CssSavingArgs args) throws Exception
        {
            // We can access the entire source document via the "Document" property.
            Assert.assertTrue(args.getDocument().getOriginalFileName().endsWith("Rendering.docx"));

            args.CssStream = new FileStream(mCssTextFileName, FileMode.CREATE);
            args.isExportNeeded(mIsExportNeeded);
            args.setKeepCssStreamOpen(mKeepCssStreamOpen);

            Assert.True(args.CssStream.CanWrite);
        }

        private /*final*/ String mCssTextFileName;
        private /*final*/ boolean mIsExportNeeded;
        private /*final*/ boolean mKeepCssStreamOpen;
    }
    //ExEnd
}
