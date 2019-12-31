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
import com.aspose.words.PdfSaveOptions;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.File;
import com.aspose.words.PdfTextCompression;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.SeekOrigin;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.words.XpsSaveOptions;
import com.aspose.words.Section;
import com.aspose.words.MultiplePagesType;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.TiffCompression;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.Drawing.msSize;
import com.aspose.ms.System.Drawing.msSizeF;
import com.aspose.words.FontSourceBase;
import com.aspose.words.FontSettings;
import java.util.ArrayList;
import com.aspose.ms.System.Collections.msArrayList;
import com.aspose.words.FolderFontSource;
import org.testng.Assert;
import com.aspose.words.SystemFontSource;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.LoadOptions;
import com.aspose.words.WarningType;
import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningInfo;
import com.aspose.words.WarningInfoCollection;
import com.aspose.words.PdfFontEmbeddingMode;
import com.aspose.words.PdfEncryptionDetails;
import com.aspose.words.PdfEncryptionAlgorithm;
import com.aspose.words.PdfPermissions;
import com.aspose.words.NumeralFormat;


@Test
public class ExRendering extends ApiExampleBase
{
    @Test
    public void saveToPdfDefault() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String)
        //ExSummary:Converts a whole document to PDF using default options.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        doc.save(getArtifactsDir() + "Rendering.SaveToPdfDefault.pdf");
        //ExEnd
    }

    @Test
    public void saveToPdfWithOutline() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:PdfSaveOptions
        //ExFor:OutlineOptions.HeadingsOutlineLevels
        //ExFor:OutlineOptions.ExpandedOutlineLevels
        //ExSummary:Converts a whole document to PDF with three levels in the document outline.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        PdfSaveOptions options = new PdfSaveOptions();
        options.getOutlineOptions().setHeadingsOutlineLevels(3);
        options.getOutlineOptions().setExpandedOutlineLevels(1);

        doc.save(getArtifactsDir() + "Rendering.SaveToPdfWithOutline.pdf", options);
        //ExEnd
    }

    @Test
    public void saveToPdfStreamOnePage() throws Exception
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.PageIndex
        //ExFor:FixedPageSaveOptions.PageCount
        //ExFor:Document.Save(Stream, SaveOptions)
        //ExSummary:Converts just one page (third page in this example) of the document to PDF.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        Stream stream = File.create(getArtifactsDir() + "Rendering.SaveToPdfStreamOnePage.pdf");
        try /*JAVA: was using*/
        {
            PdfSaveOptions options = new PdfSaveOptions();
            options.setPageIndex(2);
            options.setPageCount(1);
            doc.save(stream, options);
        }
        finally { if (stream != null) stream.close(); }

        //ExEnd
    }

    @Test
    public void saveToPdfNoCompression() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:PdfSaveOptions.TextCompression
        //ExFor:PdfTextCompression
        //ExSummary:Saves a document to PDF without compression.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setTextCompression(PdfTextCompression.NONE);

        doc.save(getArtifactsDir() + "Rendering.SaveToPdfNoCompression.pdf", options);
        //ExEnd
    }

    @Test
    public void saveAsPdf() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.PreserveFormFields
        //ExFor:Document.Save(String)
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExFor:Document.Save(String, SaveOptions)
        //ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
        // Open the document
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Option 1: Save document to file in the PDF format with default options
        doc.save(getArtifactsDir() + "Rendering.PdfDefaultOptions.pdf");

        // Option 2: Save the document to stream in the PDF format with default options
        MemoryStream stream = new MemoryStream();
        doc.save(stream, SaveFormat.PDF);
        // Rewind the stream position back to the beginning, ready for use
        stream.seek(0, SeekOrigin.BEGIN);

        // Option 3: Save document to the PDF format with specified options
        // Render the first page only and preserve form fields as usable controls and not as plain text
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.setPageIndex(0);
        pdfOptions.setPageCount(1);
        pdfOptions.setPreserveFormFields(true);
        doc.save(getArtifactsDir() + "Rendering.PdfCustomOptions.pdf", pdfOptions);
        //ExEnd
    }

    @Test
    public void saveAsXps() throws Exception
    {
        //ExStart
        //ExFor:XpsSaveOptions
        //ExFor:XpsSaveOptions.#ctor
        //ExFor:XpsSaveOptions.OutlineOptions
        //ExFor:XpsSaveOptions.SaveFormat
        //ExFor:Document.Save(String)
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExFor:Document.Save(String, SaveOptions)
        //ExSummary:Shows how to save a document to the XPS format in different ways.
        // Open the document
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Save document to file in the XPS format with default options
        doc.save(getArtifactsDir() + "Rendering.SaveAsXps.DefaultOptions.xps");

        // Save document to stream in the XPS format with default options
        FileStream docStream = new FileStream(getArtifactsDir() + "Rendering.SaveAsXps.FromStream.xps", FileMode.CREATE);
        doc.save(docStream, SaveFormat.XPS);
        docStream.close();

        // Save document to file in the XPS format with specified options
        // Render 3 pages starting from page 2; pages 2, 3 and 4
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        xpsOptions.setSaveFormat(SaveFormat.XPS);
        xpsOptions.setPageIndex(1);
        xpsOptions.setPageCount(3);

        // All paragraphs in the "Heading 1" style will be included in the outline but "Heading 2" and onwards won't
        xpsOptions.getOutlineOptions().setHeadingsOutlineLevels(1);

        doc.save(getArtifactsDir() + "Rendering.SaveAsXps.PartialDocument.xps", xpsOptions);
        //ExEnd
    }

    @Test
    public void saveAsXpsBookFold() throws Exception
    {
        //ExStart
        //ExFor:XpsSaveOptions.#ctor(SaveFormat)
        //ExFor:XpsSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the XPS format in the form of a book fold.
        // Open a document with multiple paragraphs
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Configure both page setup and XpsSaveOptions to create a book fold
        for (Section s : (Iterable<Section>) doc.getSections())
        {
            s.getPageSetup().setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);
        }

        XpsSaveOptions xpsOptions = new XpsSaveOptions(SaveFormat.XPS);
        xpsOptions.setUseBookFoldPrintingSettings(true);

        // In order to make a booklet, we will need to print this document, stack the pages
        // in the order they come out of the printer and then fold down the middle
        doc.save(getArtifactsDir() + "Rendering.SaveAsXpsBookFold.xps", xpsOptions);
        //ExEnd
    }

    @Test
    public void saveAsImage() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions.#ctor
        //ExFor:Document.Save(String)
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExFor:Document.Save(String, SaveOptions)
        //ExSummary:Shows how to save a document to the JPEG format using the Save method and the ImageSaveOptions class.
        // Open the document
        Document doc = new Document(getMyDir() + "Rendering.doc");
        // Save as a JPEG image file with default options
        doc.save(getArtifactsDir() + "Rendering.JpegDefaultOptions.jpg");

        // Save document to stream as a JPEG with default options
        MemoryStream docStream = new MemoryStream();
        doc.save(docStream, SaveFormat.JPEG);
        // Rewind the stream position back to the beginning, ready for use
        docStream.seek(0, SeekOrigin.BEGIN);

        // Save document to a JPEG image with specified options
        // Render the third page only and set the JPEG quality to 80%
        // In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor 
        // to signal what type of image to save as
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);
        imageOptions.setPageIndex(2);
        imageOptions.setPageCount(1);
        imageOptions.setJpegQuality(80);
        doc.save(getArtifactsDir() + "Rendering.JpegCustomOptions.jpg", imageOptions);
        //ExEnd
    }

    @Test (groups = "SkipMono")
    public void saveToTiffDefault() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String)
        //ExSummary:Converts a whole document into a multipage TIFF file using default options.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        doc.save(getArtifactsDir() + "Rendering.SaveToTiffDefault.tiff");
        //ExEnd
    }

    @Test (groups = "SkipMono")
    public void saveToTiffCompression() throws Exception
    {
        //ExStart
        //ExFor:TiffCompression
        //ExFor:ImageSaveOptions.TiffCompression
        //ExFor:ImageSaveOptions.PageIndex
        //ExFor:ImageSaveOptions.PageCount
        //ExFor:Document.Save(String, SaveOptions)
        //ExSummary:Converts a page of a Word document into a TIFF image and uses the CCITT compression.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.TIFF);
        {
            options.setTiffCompression(TiffCompression.CCITT_3);
            options.setPageIndex(0);
            options.setPageCount(1);
        }

        doc.save(getArtifactsDir() + "Rendering.SaveToTiffCompression.tiff", options);
        //ExEnd
    }

    @Test
    public void saveToImageResolution() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.Resolution
        //ExSummary:Renders a page of a Word document into a PNG image at a specific resolution.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
        {
            options.setResolution(300f);
            options.setPageCount(1);
        }

        doc.save(getArtifactsDir() + "Rendering.SaveToImageResolution.png", options);
        //ExEnd
    }

    @Test (groups = "SkipMono")
    public void saveToEmf() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExSummary:Converts every page of a DOC file into a separate scalable EMF file.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.EMF); { options.setPageCount(1); }

        for (int i = 0; i < doc.getPageCount(); i++)
        {
            options.setPageIndex(i);
            doc.save(getArtifactsDir() + "Rendering.SaveToEmf." + i + ".emf", options);
        }
        //ExEnd
    }

    @Test
    public void saveToImageJpegQuality() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.JpegQuality
        //ExSummary:Converts a page of a Word document into JPEG images of different qualities.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.JPEG);

        // Try worst quality
        saveOptions.setJpegQuality(0);
        doc.save(getArtifactsDir() + "Rendering.SaveToImageJpegQuality0.jpeg", saveOptions);

        // Try best quality
        saveOptions.setJpegQuality(100);
        doc.save(getArtifactsDir() + "Rendering.SaveToImageJpegQuality100.jpeg", saveOptions);
        //ExEnd
    }

    @Test
    public void saveToImagePaperColor() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.PaperColor
        //ExSummary:Renders a page of a Word document into an image with transparent or colored background.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.PNG);

        imgOptions.setPaperColor(msColor.getTransparent());
        doc.save(getArtifactsDir() + "Rendering.SaveToImagePaperColorTransparent.png", imgOptions);

        imgOptions.setPaperColor(msColor.getLightCoral());
        doc.save(getArtifactsDir() + "Rendering.SaveToImagePaperColorCoral.png", imgOptions);
        //ExEnd
    }

                @Test
    public void saveToImageStreamNetStandard2() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExSummary:Saves a document page as a BMP image into a stream (.NetStandard 2.0).
        Document doc = new Document(getMyDir() + "Rendering.doc");

        MemoryStream stream = new MemoryStream();
        doc.save(stream, SaveFormat.BMP);

        // Rewind the stream and create a .NET image from it
        stream.setPosition(0);
        
        // Read the stream back into an image
        SkiaSharp.SKBitmap image = SkiaSharp.SKBitmap.Decode(stream);
        //ExEnd
    }

    @Test
    public void renderToSizeNetStandard2() throws Exception
    {
        //ExStart
        //ExFor:Document.RenderToSize
        //ExSummary:Render to a bitmap at a specified location and size (.NetStandard 2.0).
        Document doc = new Document(getMyDir() + "Rendering.doc");
        
        SKBitmap bitmap = new SKBitmap(700, 700);
        try /*JAVA: was using*/
        {
            // User has some sort of a Graphics object. In this case created from a bitmap
            SKCanvas canvas = new SKCanvas(bitmap);
            try /*JAVA: was using*/
            {
                // Apply scale transform
                canvas.Scale(70);

                // The output should be offset 0.5" from the edge and rotated
                canvas.Translate(0.5f, 0.5f);
                canvas.RotateDegrees(10);

                // This is our test rectangle
                SKRect rect = new SKRect(0f, 0f, 3f, 3f);
                canvas.DrawRect(rect, new SKPaint();
                {
                    .setColor(SKColors.Black);
                    .setStyle(SKPaintStyle.Stroke);
                    .setStrokeWidth(3f / 72f);
                });

                // User specifies (in world coordinates) where on the Graphics to render and what size
                float returnedScale = doc.RenderToSize(0, canvas, 0f, 0f, 3f, 3f);

                msConsole.writeLine("The image was rendered at {0:P0} zoom.", returnedScale);

                // One more example, this time in millimeters
                canvas.ResetMatrix();

                // Apply scale transform
                canvas.Scale(5);

                // Move the origin 10mm 
                canvas.Translate(10, 10);

                // This is our test rectangle
                rect = new SKRect(0, 0, 50, 100);
                rect.Offset(90, 10);
                canvas.DrawRect(rect, new SKPaint();
                {
                    .setColor(SKColors.Black);
                    .setStyle(SKPaintStyle.Stroke);
                    .setStrokeWidth(1);
                });
                
                // User specifies (in world coordinates) where on the Graphics to render and what size
                doc.RenderToSize(0, canvas, 90, 10, 50, 100);

                SKFileWStream fs = new SKFileWStream(getArtifactsDir() + "Rendering.RenderToSize.png");
                try /*JAVA: was using*/
                {
                    bitmap.PeekPixels().Encode(fs, SKEncodedImageFormat.Png, 100);
                }
                finally { if (fs != null) fs.close(); }
            }
            finally { if (canvas != null) canvas.close(); }
        }
        finally { if (bitmap != null) bitmap.close(); }            
        //ExEnd
    }

    @Test
    public void createThumbnailsNetStandard2() throws Exception
    {
        //ExStart
        //ExFor:Document.RenderToScale
        //ExSummary:Renders individual pages to graphics to create one image with thumbnails of all pages (.NetStandard 2.0).
        // The user opens or builds a document
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // This defines the number of columns to display the thumbnails in
        final int THUMB_COLUMNS = 2;

        // Calculate the required number of rows for thumbnails
        // We can now get the number of pages in the document
        int thumbRows = Math.DivRem(doc.getPageCount(), THUMB_COLUMNS, /*out*/ int remainder);
        if (remainder > 0)
            thumbRows++;

        // Lets say I want thumbnails to be of this zoom
        final float SCALE = 0.25f;

        // For simplicity lets pretend all pages in the document are of the same size, 
        // so we can use the size of the first page to calculate the size of the thumbnail
        /*Size*/long thumbSize = doc.getPageInfo(0).getSizeInPixelsInternal(SCALE, 96f);

        // Calculate the size of the image that will contain all the thumbnails
        int imgWidth = msSize.getWidth(thumbSize) * THUMB_COLUMNS;
        int imgHeight = msSize.getHeight(thumbSize) * thumbRows;

        SKBitmap bitmap = new SKBitmap(imgWidth, imgHeight);
        try /*JAVA: was using*/
        {
            // The user has to provides a Graphics object to draw on
            // The Graphics object can be created from a bitmap, from a metafile, printer or window
            SKCanvas canvas = new SKCanvas(bitmap);
            try /*JAVA: was using*/
            {
                // Fill the "paper" with white, otherwise it will be transparent
                canvas.Clear(SKColors.White);

                for (int pageIndex = 0; pageIndex < doc.getPageCount(); pageIndex++)
                {
                    int rowIdx = Math.DivRem(pageIndex, THUMB_COLUMNS, /*out*/ int columnIdx);

                    // Specify where we want the thumbnail to appear
                    float thumbLeft = columnIdx * msSize.getWidth(thumbSize);
                    float thumbTop = rowIdx * msSize.getHeight(thumbSize);

                    /*SizeF*/long size = doc.RenderToScale(pageIndex, canvas, thumbLeft, thumbTop, SCALE);

                    // Draw the page rectangle
                    SKRect rect = new SKRect(0, 0, msSizeF.getWidth(size), msSizeF.getHeight(size));
                    rect.Offset(thumbLeft, thumbTop);
                    canvas.DrawRect(rect, new SKPaint();
                    {
                        .setColor(SKColors.Black);
                        .setStyle(SKPaintStyle.Stroke);
                    });
                }

                SKFileWStream fs = new SKFileWStream(getArtifactsDir() + "Rendering.Thumbnails.png");
                try /*JAVA: was using*/
                {
                    bitmap.PeekPixels().Encode(fs, SKEncodedImageFormat.Png, 100);
                }
                finally { if (fs != null) fs.close(); }
            }
            finally { if (canvas != null) canvas.close(); }
        }
        finally { if (bitmap != null) bitmap.close(); }            
        //ExEnd
    }
    
    @Test
    public void updatePageLayout() throws Exception
    {
        //ExStart
        //ExFor:StyleCollection.Item(String)
        //ExFor:SectionCollection.Item(Int32)
        //ExFor:Document.UpdatePageLayout
        //ExSummary:Shows when to request page layout of the document to be recalculated.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Saving a document to PDF or to image or printing for the first time will automatically
        // layout document pages and this information will be cached inside the document
        doc.save(getArtifactsDir() + "Rendering.UpdatePageLayout1.pdf");

        // Modify the document in any way
        doc.getStyles().get("Normal").getFont().setSize(6.0);
        doc.getSections().get(0).getPageSetup().setOrientation(com.aspose.words.Orientation.LANDSCAPE);

        // In the current version of Aspose.Words, modifying the document does not automatically rebuild 
        // the cached page layout. If you want to save to PDF or render a modified document again,
        // you need to manually request page layout to be updated
        doc.updatePageLayout();

        doc.save(getArtifactsDir() + "Rendering.UpdatePageLayout2.pdf");
        //ExEnd
    }

    @Test
    public void updateFieldsBeforeRendering() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateFields
        //ExSummary:Shows how to update all fields before rendering a document.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // This updates all fields in the document
        doc.updateFields();

        doc.save(getArtifactsDir() + "Rendering.UpdateFields.pdf");
        //ExEnd
    }

    @Test
    public void setTrueTypeFontsFolder() throws Exception
    {
        // Store the font sources currently used so we can restore them later
        FontSourceBase[] fontSources = FontSettings.getDefaultInstance().getFontsSources();

        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.SetFontsFolder(String, Boolean)
        //ExSummary:Demonstrates how to set the folder Aspose.Words uses to look for TrueType fonts during rendering or embedding of fonts.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Note that this setting will override any default font sources that are being searched by default
        // Now only these folders will be searched for fonts when rendering or embedding fonts
        // To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and 
        // FontSettings.SetFontSources instead
        FontSettings.getDefaultInstance().setFontsFolder("C:\\MyFonts\\", false);

        doc.save(getArtifactsDir() + "Rendering.SetFontsFolder.pdf");
        //ExEnd

        // Restore the original sources used to search for fonts
        FontSettings.getDefaultInstance().setFontsSources(fontSources);
    }

    @Test
    public void setFontsFoldersMultipleFolders() throws Exception
    {
        // Store the font sources currently used so we can restore them later
        FontSourceBase[] fontSources = FontSettings.getDefaultInstance().getFontsSources();

        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.SetFontsFolders(String[], Boolean)
        //ExSummary:Demonstrates how to set Aspose.Words to look in multiple folders for TrueType fonts when rendering or embedding fonts.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Note that this setting will override any default font sources that are being searched by default
        // Now only these folders will be searched for fonts when rendering or embedding fonts
        // To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and 
        // FontSettings.SetFontSources instead
        FontSettings.getDefaultInstance().setFontsFolders(new String[] { "C:\\MyFonts\\", "D:\\Misc\\Fonts\\" }, true);

        doc.save(getArtifactsDir() + "Rendering.SetFontsFolders.pdf");
        //ExEnd

        // Restore the original sources used to search for fonts
        FontSettings.getDefaultInstance().setFontsSources(fontSources);
    }

    @Test
    public void setFontsFoldersSystemAndCustomFolder() throws Exception
    {
        // Store the font sources currently used so we can restore them later
        FontSourceBase[] origFontSources = FontSettings.getDefaultInstance().getFontsSources();

        //ExStart
        //ExFor:FontSettings            
        //ExFor:FontSettings.GetFontsSources()
        //ExFor:FontSettings.SetFontsSources()
        //ExSummary:Demonstrates how to set Aspose.Words to look for TrueType fonts in system folders as well as a custom defined folder when scanning for fonts.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Retrieve the array of environment-dependent font sources that are searched by default
        // For example this will contain a "Windows\Fonts\" source on a Windows machines
        // We add this array to a new ArrayList to make adding or removing font entries much easier
        ArrayList fontSources = msArrayList.ctor(FontSettings.getDefaultInstance().getFontsSources());

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts
        FolderFontSource folderFontSource = new FolderFontSource("C:\\MyFonts\\", true);

        // Add the custom folder which contains our fonts to the list of existing font sources
        msArrayList.add(fontSources, folderFontSource);

        // Convert the ArrayList of source back into a primitive array of FontSource objects
        FontSourceBase[] updatedFontSources = (FontSourceBase[]) msArrayList.toArray(fontSources, FontSourceBase.class);

        // Apply the new set of font sources to use
        FontSettings.getDefaultInstance().setFontsSources(updatedFontSources);

        doc.save(getArtifactsDir() + "Rendering.SetFontsFolders.pdf");
        //ExEnd

        // The first source should be a system font source
        Assert.That(FontSettings.getDefaultInstance().getFontsSources()[0], Is.InstanceOf(SystemFontSource.class)); 
        // The second source should be our folder font source
        Assert.That(FontSettings.getDefaultInstance().getFontsSources()[1], Is.InstanceOf(FolderFontSource.class)); 
        
        FolderFontSource folderSource = ((FolderFontSource) FontSettings.getDefaultInstance().getFontsSources()[1]);
        msAssert.areEqual("C:\\MyFonts\\", folderSource.getFolderPath());
        Assert.assertTrue(folderSource.getScanSubfolders());

        // Restore the original sources used to search for fonts
        FontSettings.getDefaultInstance().setFontsSources(origFontSources);
    }

    @Test
    public void setSpecifyFontFolder() throws Exception
    {
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolder(getMyDir() + "MyFonts\\", false);

        // Using load options
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);

        Document doc = new Document(getMyDir() + "Rendering.doc", loadOptions);

        FolderFontSource folderSource = ((FolderFontSource) doc.getFontSettings().getFontsSources()[0]);

        msAssert.areEqual(getMyDir() + "MyFonts\\", folderSource.getFolderPath());
        Assert.assertFalse(folderSource.getScanSubfolders());
    }

    @Test
    public void setFontSubstitutes() throws Exception
    {
        //ExStart
        //ExFor:Document.FontSettings
        //ExFor:TableSubstitutionRule.SetSubstitutes(String, String[])
        //ExSummary:Shows how to define alternative fonts if original does not exist
        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getTableSubstitution().setSubstitutes("Times New Roman", new String[] { "Slab", "Arvo" });
        //ExEnd
        Document doc = new Document(getMyDir() + "Rendering.doc");
        doc.setFontSettings(fontSettings);

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        // Check that font source are default
        FontSourceBase[] fontSource = doc.getFontSettings().getFontsSources();
        msAssert.areEqual("SystemFonts", FontSourceType.toString(fontSource[0].getType()));

        msAssert.areEqual("Times New Roman", doc.getFontSettings().getSubstitutionSettings().getDefaultFontSubstitution().getDefaultFontName());

        String[] alternativeFonts = doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Times New Roman").ToArray();
        msAssert.areEqual(new String[] { "Slab", "Arvo" }, alternativeFonts);
    }

    @Test
    public void setSpecifyFontFolders() throws Exception
    {
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolders(new String[] { getMyDir() + "MyFonts\\", "C:\\Windows\\Fonts\\" }, true);

        // Using load options
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);
        Document doc = new Document(getMyDir() + "Rendering.doc", loadOptions);

        FolderFontSource folderSource = ((FolderFontSource) doc.getFontSettings().getFontsSources()[0]);
        msAssert.areEqual(getMyDir() + "MyFonts\\", folderSource.getFolderPath());
        Assert.assertTrue(folderSource.getScanSubfolders());

        folderSource = ((FolderFontSource) doc.getFontSettings().getFontsSources()[1]);
        msAssert.areEqual("C:\\Windows\\Fonts\\", folderSource.getFolderPath());
        Assert.assertTrue(folderSource.getScanSubfolders());
    }

    @Test
    public void addFontSubstitutes() throws Exception
    {
        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getTableSubstitution().setSubstitutes("Slab", new String[] { "Times New Roman", "Arial" });
        fontSettings.getSubstitutionSettings().getTableSubstitution().addSubstitutes("Arvo", new String[] { "Open Sans", "Arial" });

        Document doc = new Document(getMyDir() + "Rendering.doc");
        doc.setFontSettings(fontSettings);

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        String[] alternativeFonts = doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Slab").ToArray();
        msAssert.areEqual(new String[] { "Times New Roman", "Arial" }, alternativeFonts);

        alternativeFonts = doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Arvo").ToArray();
        msAssert.areEqual(new String[] { "Open Sans", "Arial" }, alternativeFonts);
    }

    @Test
    public void setDefaultFontName() throws Exception
    {
        //ExStart
        //ExFor:DefaultFontSubstitutionRule.DefaultFontName
        //ExSummary:Demonstrates how to specify what font to substitute for a missing font during rendering.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead
        FontSettings.getDefaultInstance().getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial Unicode MS");

        // Now the set default font is used in place of any missing fonts during any rendering calls
        doc.save(getArtifactsDir() + "Rendering.SetDefaultFont.pdf");
        doc.save(getArtifactsDir() + "Rendering.SetDefaultFont.xps");
        //ExEnd
    }

    @Test
    public void recieveFontSubstitutionUpdatePageLayout() throws Exception
    {
        // Store the font sources currently used so we can restore them later
        FontSourceBase[] origFontSources = FontSettings.getDefaultInstance().getFontsSources();

        // Load the document to render
        Document doc = new Document(getMyDir() + "Document.doc");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        // We can choose the default font to use in the case of any missing fonts
        FontSettings.getDefaultInstance().getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");

        // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default 
        // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback
        FontSettings.getDefaultInstance().setFontsFolder("", false);

        // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occurred during rendering
        // are stored until the document save and then sent to the appropriate WarningCallback
        doc.updatePageLayout();

        // Even though the document was rendered previously, any save warnings are notified to the user during document save
        doc.save(getArtifactsDir() + "Rendering.FontsNotificationUpdatePageLayout.pdf");
        
        Assert.That(callback.FontWarnings.getCount(), Is.GreaterThan(0));
        Assert.assertTrue(callback.FontWarnings.get(0).getWarningType() == WarningType.FONT_SUBSTITUTION);
        Assert.assertTrue(callback.FontWarnings.get(0).getDescription().contains("has not been found"));

        // Restore default fonts
        FontSettings.getDefaultInstance().setFontsSources(origFontSources);
    }

    public static class HandleDocumentWarnings implements IWarningCallback
    {
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        public void warning(WarningInfo info)
        {
            // We are only interested in fonts being substituted
            if (info.getWarningType() == WarningType.FONT_SUBSTITUTION)
            {
                msConsole.writeLine("Font substitution: " + info.getDescription());
                FontWarnings.warning(info); //ExSkip
            }
        }

        public WarningInfoCollection FontWarnings = new WarningInfoCollection(); //ExSkip
    }

    @Test
    public void embedFullFontsInPdf() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.#ctor
        //ExFor:PdfSaveOptions.EmbedFullFonts
        //ExSummary:Demonstrates how to set Aspose.Words to embed full fonts in the output PDF document.
        // Load the document to render
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true
        // The property below can be changed each time a document is rendered
        PdfSaveOptions options = new PdfSaveOptions();
        options.setEmbedFullFonts(true);

        // The output PDF will be embedded with all fonts found in the document
        doc.save(getArtifactsDir() + "Rendering.EmbedFullFonts.pdf");
        //ExEnd
    }

    @Test
    public void subsetFontsInPdf() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.EmbedFullFonts
        //ExSummary:Demonstrates how to set Aspose.Words to subset fonts in the output PDF.
        // Load the document to render
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false
        PdfSaveOptions options = new PdfSaveOptions();
        options.setEmbedFullFonts(false);

        // The output PDF will contain subsets of the fonts in the document
        // Only the glyphs used in the document are included in the PDF fonts
        doc.save(getArtifactsDir() + "Rendering.SubsetFonts.pdf");
        //ExEnd
    }

    @Test
    public void disableEmbeddingStandardWindowsFonts() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.FontEmbeddingMode
        //ExFor:PdfFontEmbeddingMode
        //ExSummary:Shows how to set Aspose.Words to skip embedding Arial and Times New Roman fonts into a PDF document.
        // Load the document to render
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false
        PdfSaveOptions options = new PdfSaveOptions();
        options.setFontEmbeddingMode(PdfFontEmbeddingMode.EMBED_NONE);

        // The output PDF will be saved without embedding standard windows fonts
        doc.save(getArtifactsDir() + "Rendering.DisableEmbedWindowsFonts.pdf");
        //ExEnd
    }

    @Test
    public void disableEmbeddingCoreFonts() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.UseCoreFonts
        //ExSummary:Shows how to set Aspose.Words to avoid embedding core fonts and let the reader substitute PDF Type 1 fonts instead.
        // Load the document to render
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // To disable embedding of core fonts and substitute PDF type 1 fonts set UseCoreFonts to true
        PdfSaveOptions options = new PdfSaveOptions();
        options.setUseCoreFonts(true);

        // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        doc.save(getArtifactsDir() + "Rendering.DisableEmbedWindowsFonts.pdf");
        //ExEnd
    }

    @Test
    public void setPdfEncryptionPermissions() throws Exception
    {
        //ExStart
        //ExFor:PdfEncryptionDetails.#ctor
        //ExFor:PdfSaveOptions.EncryptionDetails
        //ExFor:PdfEncryptionDetails.Permissions
        //ExFor:PdfEncryptionDetails.EncryptionAlgorithm
        //ExFor:PdfEncryptionDetails.OwnerPassword
        //ExFor:PdfEncryptionDetails.UserPassword
        //ExFor:PdfEncryptionAlgorithm
        //ExFor:PdfPermissions
        //ExFor:PdfEncryptionDetails
        //ExSummary:Demonstrates how to set permissions on a PDF document generated by Aspose.Words.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Create encryption details and set owner password
        PdfEncryptionDetails encryptionDetails =
            new PdfEncryptionDetails("password", "", PdfEncryptionAlgorithm.RC_4_128);

        // Start by disallowing all permissions
        encryptionDetails.setPermissions(PdfPermissions.DISALLOW_ALL);

        // Extend permissions to allow editing or modifying annotations
        encryptionDetails.setPermissions(PdfPermissions.MODIFY_ANNOTATIONS | PdfPermissions.DOCUMENT_ASSEMBLY);
        saveOptions.setEncryptionDetails(encryptionDetails);

        // Render the document to PDF format with the specified permissions
        doc.save(getArtifactsDir() + "Rendering.SpecifyPermissions.pdf", saveOptions);
        //ExEnd
    }

    @Test
    public void setPdfNumeralFormat() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.NumeralFormat.doc");
        //ExStart
        //ExFor:FixedPageSaveOptions.NumeralFormat
        //ExFor:NumeralFormat
        //ExSummary:Demonstrates how to set the numeral format used when saving to PDF.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setNumeralFormat(NumeralFormat.CONTEXT);
        //ExEnd

        doc.save(getArtifactsDir() + "Rendering.NumeralFormat.pdf", options);
    }
}
