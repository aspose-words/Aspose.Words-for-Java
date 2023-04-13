package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.concurrent.TimeUnit;

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
    //ExSummary:Shows how to print the filenames of linked images created while converting a document to flow-form .xaml.
    @Test //ExSkip
    public void imageFolder() throws Exception {
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
        new File(options.getImagesFolderAlias()).mkdir();

        doc.save(getArtifactsDir() + "XamlFlowSaveOptions.ImageFolder.xaml", options);

        for (String resource : callback.getResources())
            System.out.println("{callback.ImagesFolderAlias}/{resource}");
        testImageFolder(callback); //ExSkip
    }

    /// <summary>
    /// Counts and prints filenames of images while their parent document is converted to flow-form .xaml.
    /// </summary>
    private static class ImageUriPrinter implements IImageSavingCallback {
        public ImageUriPrinter(String imagesFolderAlias) {
            mImagesFolderAlias = imagesFolderAlias;
            mResources = new ArrayList<String>();
        }

        public void imageSaving(ImageSavingArgs args) throws Exception {
            getResources().add(args.getImageFileName());

            // If we specified an image folder alias, we would also need
            // to redirect each stream to put its image in the alias folder.
            args.setImageStream(new FileOutputStream(MessageFormat.format("{0}/{1}", mImagesFolderAlias, args.getImageFileName())));
            args.setKeepImageStreamOpen(false);
        }

        public String getImagesFolderAlias() {
            return mImagesFolderAlias;
        }

        private final String mImagesFolderAlias;

        public ArrayList<String> getResources() {
            return mResources;
        }

        private final ArrayList<String> mResources;
    }
    //ExEnd

    private void testImageFolder(ImageUriPrinter callback) {
        Assert.assertEquals(9, callback.getResources().size());
        for (String resource : callback.getResources())
            Assert.assertTrue(new File(MessageFormat.format("{0}/{1}", callback.getImagesFolderAlias(), resource)).exists());
    }

    @Test (dataProvider = "progressCallbackDataProvider")
    //ExStart
    //ExFor:SaveOptions.ProgressCallback
    //ExFor:IDocumentSavingCallback.Notify(DocumentSavingArgs)
    //ExFor:DocumentSavingArgs.EstimatedProgress
    //ExSummary:Shows how to manage a document while saving to xamlflow.
    public void progressCallback(int saveFormat, String ext) throws Exception
    {
        Document doc = new Document(getMyDir() + "Big document.docx");

        // Following formats are supported: XamlFlow, XamlFlowPack.
        XamlFlowSaveOptions saveOptions = new XamlFlowSaveOptions(saveFormat);
        {
            saveOptions.setProgressCallback(new SavingProgressCallback());
        }

        try {
            doc.save(getArtifactsDir() + MessageFormat.format("XamlFlowSaveOptions.ProgressCallback.{0}", ext), saveOptions);
        }
        catch (IllegalStateException exception) {
            Assert.assertTrue(exception.getMessage().contains("EstimatedProgress"));
        }
    }

    @DataProvider(name = "progressCallbackDataProvider") //ExSkip
    public static Object[][] progressCallbackDataProvider() throws Exception
    {
        return new Object[][]
                {
                        {SaveFormat.XAML_FLOW,  "xamlflow"},
                        {SaveFormat.XAML_FLOW_PACK,  "xamlflowpack"},
                };
    }

    /// <summary>
    /// Saving progress callback. Cancel a document saving after the "MaxDuration" seconds.
    /// </summary>
    public static class SavingProgressCallback implements IDocumentSavingCallback
    {
        /// <summary>
        /// Ctr.
        /// </summary>
        public SavingProgressCallback()
        {
            mSavingStartedAt = new Date();
        }

        /// <summary>
        /// Callback method which called during document saving.
        /// </summary>
        /// <param name="args">Saving arguments.</param>
        public void notify(DocumentSavingArgs args)
        {
            Date canceledAt = new Date();
            long diff = canceledAt.getTime() - mSavingStartedAt.getTime();
            long ellapsedSeconds = TimeUnit.MILLISECONDS.toSeconds(diff);

            if (ellapsedSeconds > MAX_DURATION)
                throw new IllegalStateException(MessageFormat.format("EstimatedProgress = {0}; CanceledAt = {1}", args.getEstimatedProgress(), canceledAt));
        }

        /// <summary>
        /// Date and time when document saving is started.
        /// </summary>
        private Date mSavingStartedAt;

        /// <summary>
        /// Maximum allowed duration in sec.
        /// </summary>
        private static final double MAX_DURATION = 0.01d;
    }
    //ExEnd
}