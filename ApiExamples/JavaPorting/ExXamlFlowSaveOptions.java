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
import com.aspose.words.XamlFlowSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.Directory;
import com.aspose.words.IImageSavingCallback;
import com.aspose.words.ImageSavingArgs;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;


@Test
public class ExXamlFlowSaveOptions extends ApiExampleBase
{
    //ExStart
    //ExFor:XamlFlowSaveOptions
    //ExFor:XamlFlowSaveOptions.#ctor
    //ExFor:XamlFlowSaveOptions.#ctor(SaveFormat)
    //ExFor:XamlFlowSaveOptions.ImageSavingCallback
    //ExFor:XamlFlowSaveOptions.ImagesFolder
    //ExFor:XamlFlowSaveOptions.ImagesFolderAlias
    //ExFor:XamlFlowSaveOptions.SaveFormat
    //ExSummary:Shows how to print the filenames of linked images created during conversion of a document to flow-form .xaml.
    @Test //ExSkip
    public void imageFolder() throws Exception
    {
        // Open a document which contains images
        Document doc = new Document(getMyDir() + "Rendering.docx");

        XamlFlowSaveOptions options = new XamlFlowSaveOptions();
        {
            options.setSaveFormat(SaveFormat.XAML_FLOW);
            options.setImagesFolder(getArtifactsDir() + "XamlFlowImageFolder");
            options.setImagesFolderAlias(getArtifactsDir() + "XamlFlowImageFolderAlias");
            options.setImageSavingCallback(new ImageUriPrinter(getArtifactsDir() + "XamlFlowImageFolderAlias"));
        }

        // A folder specified by ImagesFolderAlias will contain the images instead of ImagesFolder
        // We must ensure the folder exists before the streams can put their images into it
        Directory.createDirectory(options.getImagesFolderAlias());

        doc.save(getArtifactsDir() + "XamlFlowSaveOptions.ImageFolder.xaml", options);
    }

    /// <summary>
    /// Counts and prints filenames of images while their parent document is converted to flow-form .xaml.
    /// </summary>
    private static class ImageUriPrinter implements IImageSavingCallback
    {
        public ImageUriPrinter(String imagesFolderAlias)
        {
            mImagesFolderAlias = imagesFolderAlias;
        }

        public void /*IImageSavingCallback.*/imageSaving(ImageSavingArgs args) throws Exception
        {
            msConsole.writeLine($"Image #{++mSavedImageCount} \"{args.ImageFileName}\"");

            // If we specified a ImagesFolderAlias we will also need to redirect each stream to put its image in that folder
            args.ImageStream = new FileStream($"{mImagesFolderAlias}/{args.ImageFileName}", FileMode.CREATE);
            args.setKeepImageStreamOpen(false);
        }

        private int mSavedImageCount;
        private /*final*/ String mImagesFolderAlias;
    }
    //ExEnd
}
