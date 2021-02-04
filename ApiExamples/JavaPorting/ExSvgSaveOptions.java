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
import com.aspose.words.SvgSaveOptions;
import com.aspose.words.SvgTextOutputMode;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.Directory;
import com.aspose.words.IResourceSavingCallback;
import com.aspose.words.ResourceSavingArgs;
import com.aspose.ms.System.msConsole;


@Test
public class ExSvgSaveOptions extends ApiExampleBase
{
    @Test
    public void saveLikeImage() throws Exception
    {
        //ExStart
        //ExFor:SvgSaveOptions.FitToViewPort
        //ExFor:SvgSaveOptions.ShowPageBorder
        //ExFor:SvgSaveOptions.TextOutputMode
        //ExFor:SvgTextOutputMode
        //ExSummary:Shows how to mimic the properties of images when converting a .docx document to .svg.
        Document doc = new Document(getMyDir() + "Document.docx");

        // Configure the SvgSaveOptions object to save with no page borders or selectable text.
        SvgSaveOptions options = new SvgSaveOptions();
        {
            options.setFitToViewPort(true);
            options.setShowPageBorder(false);
            options.setTextOutputMode(SvgTextOutputMode.USE_PLACED_GLYPHS);
        }

        doc.save(getArtifactsDir() + "SvgSaveOptions.SaveLikeImage.svg", options);
        //ExEnd
    }

    //ExStart
    //ExFor:SvgSaveOptions
    //ExFor:SvgSaveOptions.ExportEmbeddedImages
    //ExFor:SvgSaveOptions.ResourceSavingCallback
    //ExFor:SvgSaveOptions.ResourcesFolder
    //ExFor:SvgSaveOptions.ResourcesFolderAlias
    //ExFor:SvgSaveOptions.SaveFormat
    //ExSummary:Shows how to manipulate and print the URIs of linked resources created while converting a document to .svg.
    @Test //ExSkip
    public void svgResourceFolder() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        SvgSaveOptions options = new SvgSaveOptions();
        {
            options.setSaveFormat(SaveFormat.SVG);
            options.setExportEmbeddedImages(false);
            options.setResourcesFolder(getArtifactsDir() + "SvgResourceFolder");
            options.setResourcesFolderAlias(getArtifactsDir() + "SvgResourceFolderAlias");
            options.setShowPageBorder(false);

            options.setResourceSavingCallback(new ResourceUriPrinter());
        }

        Directory.createDirectory(options.getResourcesFolderAlias());

        doc.save(getArtifactsDir() + "SvgSaveOptions.SvgResourceFolder.svg", options);
    }

    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to .svg.
    /// </summary>
    private static class ResourceUriPrinter implements IResourceSavingCallback
    {
        public void /*IResourceSavingCallback.*/resourceSaving(ResourceSavingArgs args)
        {
            System.out.println("Resource #{++mSavedResourceCount} \"{args.ResourceFileName}\"");
            System.out.println("\t" + args.getResourceFileUri());
        }

        private int mSavedResourceCount;
    }
    //ExEnd
}
