// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.msString;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.IPageSavingCallback;
import com.aspose.words.PageSavingArgs;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.DocumentSplitCriteria;
import com.aspose.words.IDocumentPartSavingCallback;
import com.aspose.words.DocumentPartSavingArgs;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
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
        htmlFixedSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        PsSaveOptions psSaveOptions = new PsSaveOptions();
        psSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        SvgSaveOptions svgSaveOptions = new SvgSaveOptions();
        svgSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        XamlFixedSaveOptions xamlFixedSaveOptions = new XamlFixedSaveOptions();
        xamlFixedSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        XpsSaveOptions xpsSaveOptions = new XpsSaveOptions();
        xpsSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());
    }

    @Test
    public void pageFileNameSavingCallback() throws Exception
    {
        //ExStart
        //ExFor:IPageSavingCallback
        //ExFor:PageSavingArgs
        //ExFor:PageSavingArgs.PageFileName
        //ExFor:FixedPageSaveOptions.PageSavingCallback
        //ExSummary:Shows how separate pages are saved when a document is exported to fixed page format.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        HtmlFixedSaveOptions htmlFixedSaveOptions =
            new HtmlFixedSaveOptions(); { htmlFixedSaveOptions.setPageIndex(0); htmlFixedSaveOptions.setPageCount(doc.getPageCount()); }
        htmlFixedSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        doc.save(getArtifactsDir() + "Rendering.html", htmlFixedSaveOptions);

        String[] filePaths = Directory.getFiles(getArtifactsDir() + "", "Page_*.html");

        for (int i = 0; i < doc.getPageCount(); i++)
        {
            String file = msString.format(getArtifactsDir() + "Page_{0}.html", i);
            msAssert.areEqual(file, filePaths[i]);//ExSkip
        }
    }

    /// <summary>
    /// Custom PageFileName is specified.
    /// </summary>
    private static class CustomPageFileNamePageSavingCallback implements IPageSavingCallback
    {
        public void pageSaving(PageSavingArgs args)
        {
            // Specify name of the output file for the current page.
            args.setPageFileName(msString.format(getArtifactsDir() + "Page_{0}.html", args.getPageIndex()));
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
    //ExFor:HtmlSaveOptions.ImageSavingCallback
    //ExSummary:Shows how split a document into parts and save them.
    @Test //ExSkip
    public void documentParts() throws Exception
    {
        // Open a document to be converted to html
        Document doc = new Document(getMyDir() + "Rendering.doc");
        String outFileName = "SavingCallback.DocumentParts.html";

        // We can use an appropriate SaveOptions subclass to customize the conversion process
        HtmlSaveOptions options = new HtmlSaveOptions();

        // We can use it to split a document into smaller parts, in this instance split by section breaks
        // Each part will be saved into a separate file, creating many files during the conversion process instead of just one
        options.setDocumentSplitCriteria(DocumentSplitCriteria.SECTION_BREAK);

        // We can set a callback to name each document part file ourselves
        options.setDocumentPartSavingCallback(new SavedDocumentPartRename(outFileName, options.getDocumentSplitCriteria()));

        // If we convert a document that contains images into html, we will end up with one html file which links to several images
        // Each image will be in the form of a file in the local file system
        // There is also a callback that can customize the name and file system location of each image
        options.setImageSavingCallback(new SavedImageRename(outFileName));

        // The DocumentPartSaving() and ImageSaving() methods of our callbacks will be run at this time
        doc.save(getArtifactsDir() + outFileName, options);
    }

    /// <summary>
    /// Renames saved document parts that are produced when an HTML document is saved while being split according to a criteria
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
            Assert.assertTrue(args.getDocument().getOriginalFileName().endsWith("Rendering.doc"));

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

            // We can designate the filename and location of each output file either by filename
            args.setDocumentPartFileName(partFileName);

            // Or we can make a new stream and choose the location of the file at construction
            args.setDocumentPartStreamInternal(new FileStream(getArtifactsDir() + partFileName, FileMode.CREATE));
            Assert.assertTrue(args.getDocumentPartStreamInternal().canWrite());
            Assert.assertFalse(args.getKeepDocumentPartStreamOpen());
        }

        private int mCount;
        private /*final*/ String mOutFileName;
        private /*final*/ /*DocumentSplitCriteria*/int mDocumentSplitCriteria;
    }

    /// <summary>
    /// Renames saved images that are produced when an HTML document is saved 
    /// </summary>
    public static class SavedImageRename implements IImageSavingCallback
    {
        public SavedImageRename(String outFileName)
        {
            mOutFileName = outFileName;
        }

        public void /*IImageSavingCallback.*/imageSaving(ImageSavingArgs args) throws Exception
        {
            // Same filename and stream functions as above in IDocumentPartSavingCallback apply here
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
	
    //ExStart
    //ExFor:CssSavingArgs
    //ExFor:CssSavingArgs.CssStream
    //ExFor:CssSavingArgs.Document
    //ExFor:CssSavingArgs.IsExportNeeded
    //ExFor:CssSavingArgs.KeepCssStreamOpen
    //ExFor:CssStyleSheetType
    //ExFor:ICssSavingCallback
    //ExFor:ICssSavingCallback.CssSaving(CssSavingArgs)
    //ExSummary:Shows how to work with CSS stylesheets that may be created along with Html documents.
    @Test //ExSkip
    public void cssSavingCallback() throws Exception
    {
        // Open a document to be converted to html
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // If our output document will produce a CSS stylesheet, we can use an HtmlSaveOptions to control where it is saved
        HtmlSaveOptions htmlFixedSaveOptions = new HtmlSaveOptions();

        // By default, a CSS stylesheet is stored inside its HTML document, but we can have it saved to a separate file
        htmlFixedSaveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);

        // A custom ICssSavingCallback implementation can control where that stylesheet will be saved and linked to by the Html document
        htmlFixedSaveOptions.setCssSavingCallback(new CustomCssSavingCallback(getArtifactsDir() + "Rendering.CssSavingCallback.css", true, false));

        // The CssSaving() method of our callback will be called at this stage
        doc.save(getArtifactsDir() + "Rendering.CssSavingCallback.html", htmlFixedSaveOptions);
    }

    /// <summary>
    /// Designates a filename and other parameters for the saving of a CSS stylesheet
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
            // Set up the stream that will create the CSS document         
            args.CssStream = new FileStream(mCssTextFileName, FileMode.CREATE);
            Assert.True(args.CssStream.CanWrite);
            args.isExportNeeded(mIsExportNeeded);
            args.setKeepCssStreamOpen(mKeepCssStreamOpen);

            // We can also access the original document here like this
            Assert.assertTrue(args.getDocument().getOriginalFileName().endsWith("Rendering.doc"));
        }

        private /*final*/ String mCssTextFileName;
        private /*final*/ boolean mIsExportNeeded;
        private /*final*/ boolean mKeepCssStreamOpen;
    }
    //ExEnd
}
