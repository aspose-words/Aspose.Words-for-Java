// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package Examples;

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.text.MessageFormat;

@Test
public class ExXamlFlowSaveOptions extends ApiExampleBase {
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
    public void xamlFlowImageFolder() throws Exception {
        // Open a document which contains images
        Document doc = new Document(getMyDir() + "Rendering.doc");

        XamlFlowSaveOptions options = new XamlFlowSaveOptions();
        {
            options.setSaveFormat(SaveFormat.XAML_FLOW);
            options.setImagesFolder(getArtifactsDir() + "XamlFlowImageFolder");
            options.setImagesFolderAlias(getArtifactsDir() + "XamlFlowImageFolderAlias");
            options.setImageSavingCallback(new ImageUriPrinter(getArtifactsDir() + "XamlFlowImageFolderAlias"));
        }

        // A folder specified by ImagesFolderAlias will contain the images instead of ImagesFolder
        // We must ensure the folder exists before the streams can put their images into it
        new File(options.getImagesFolderAlias()).mkdir();

        doc.save(getArtifactsDir() + "XamlFlowImageFolder.xaml", options);
    }

    /// <summary>
    /// Counts and prints filenames of images while their parent document is converted to flow-form .xaml
    /// </summary>
    private static class ImageUriPrinter implements IImageSavingCallback {
        public ImageUriPrinter(String imagesFolderAlias) {
            mImagesFolderAlias = imagesFolderAlias;
        }

        public void imageSaving(ImageSavingArgs args) throws Exception {
            System.out.println(MessageFormat.format("Image #{0} \"{1}\"", ++mSavedImageCount, args.getImageFileName()));

            // If we specified a ImagesFolderAlias we will also need to redirect each stream to put its image in that folder
            args.setImageStream(new FileOutputStream(MessageFormat.format("{0}/{1}", mImagesFolderAlias, args.getImageFileName())));
            args.setKeepImageStreamOpen(false);
        }

        private int mSavedImageCount;
        private String mImagesFolderAlias;
    }
    //ExEnd
}