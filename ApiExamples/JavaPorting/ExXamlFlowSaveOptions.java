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
import com.aspose.words.XamlFlowSaveOptions;
import org.testng.Assert;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.msConsole;
import com.aspose.words.IImageSavingCallback;
import java.util.ArrayList;
import com.aspose.words.ImageSavingArgs;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.IO.File;


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
    //ExSummary:Shows how to print the filenames of linked images created while converting a document to flow-form .xaml.
    @Test //ExSkip
    public void imageFolder() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        ImageUriPrinter callback = new ImageUriPrinter(getArtifactsDir() + "XamlFlowImageFolderAlias");

        // Create a "XamlFlowSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to the XAML save format.
        XamlFlowSaveOptions options = new XamlFlowSaveOptions();

        Assert.assertEquals(SaveFormat.XAML_FLOW, options.getSaveFormat());

        // Use the "ImagesFolder" property to assign a folder in the local file system into which
        // Aspose.Words will save all the document's linked images.
        options.setImagesFolder(getArtifactsDir() + "XamlFlowImageFolder");

        // Use the "ImagesFolderAlias" property to use this folder
        // when constructing image URIs instead of the images folder's name.
        options.setImagesFolderAlias(getArtifactsDir() + "XamlFlowImageFolderAlias");

        options.setImageSavingCallback(callback);

        // A folder specified by "ImagesFolderAlias" will need to contain the resources instead of "ImagesFolder".
        // We must ensure the folder exists before the callback's streams can put their resources into it.
        Directory.createDirectory(options.getImagesFolderAlias());

        doc.save(getArtifactsDir() + "XamlFlowSaveOptions.ImageFolder.xaml", options);

        for (String resource : callback.getResources())
            System.out.println("{callback.ImagesFolderAlias}/{resource}");
        testImageFolder(callback); //ExSkip
    }

    /// <summary>
    /// Counts and prints filenames of images while their parent document is converted to flow-form .xaml.
    /// </summary>
    private static class ImageUriPrinter implements IImageSavingCallback
    {
        public ImageUriPrinter(String imagesFolderAlias)
        {
            mImagesFolderAlias = imagesFolderAlias;
            mResources = new ArrayList<String>();
        }

        public void /*IImageSavingCallback.*/imageSaving(ImageSavingArgs args) throws Exception
        {
            getResources().add(args.getImageFileName());

            // If we specified an image folder alias, we would also need
            // to redirect each stream to put its image in the alias folder.
            args.ImageStream = new FileStream($"{ImagesFolderAlias}/{args.ImageFileName}", FileMode.CREATE);
            args.setKeepImageStreamOpen(false);
        }

        public String getImagesFolderAlias() { return mImagesFolderAlias; };

        private  String mImagesFolderAlias;
        public ArrayList<String> getResources() { return mResources; };

        private ArrayList<String> mResources;
    }
    //ExEnd

    private void testImageFolder(ImageUriPrinter callback) throws Exception
    {
        Assert.assertEquals(9, callback.getResources().size());
        for (String resource : callback.getResources())
            Assert.assertTrue(File.exists($"{callback.ImagesFolderAlias}/{resource}"));
    }
}
