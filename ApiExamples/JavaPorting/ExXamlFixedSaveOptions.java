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
import com.aspose.words.XamlFixedSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.Directory;
import com.aspose.words.IResourceSavingCallback;
import com.aspose.words.ResourceSavingArgs;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;


@Test
public class ExXamlFixedSaveOptions extends ApiExampleBase
{
    //ExStart
    //ExFor:XamlFixedSaveOptions
    //ExFor:XamlFixedSaveOptions.ResourceSavingCallback
    //ExFor:XamlFixedSaveOptions.ResourcesFolder
    //ExFor:XamlFixedSaveOptions.ResourcesFolderAlias
    //ExFor:XamlFixedSaveOptions.SaveFormat
    //ExSummary:Shows how to print the URIs of linked resources created during conversion of a document to fixed-form .xaml.
    @Test //ExSkip
    public void resourceFolder() throws Exception
    {
        // Open a document which contains resources
        Document doc = new Document(getMyDir() + "Rendering.docx");

        XamlFixedSaveOptions options = new XamlFixedSaveOptions();
        {
            options.setSaveFormat(SaveFormat.XAML_FIXED);
            options.setResourcesFolder(getArtifactsDir() + "XamlFixedResourceFolder");
            options.setResourcesFolderAlias(getArtifactsDir() + "XamlFixedFolderAlias");
            options.setResourceSavingCallback(new ResourceUriPrinter());
        }

        // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder
        // We must ensure the folder exists before the streams can put their resources into it
        Directory.createDirectory(options.getResourcesFolderAlias());

        doc.save(getArtifactsDir() + "XamlFixedSaveOptions.ResourceFolder.xaml", options);
    }

    /// <summary>
    /// Counts and prints URIs of resources created during conversion to to fixed .xaml.
    /// </summary>
    private static class ResourceUriPrinter implements IResourceSavingCallback
    {
        public void /*IResourceSavingCallback.*/resourceSaving(ResourceSavingArgs args) throws Exception
        {
            // If we set a folder alias in the SaveOptions object, it will be printed here
            System.out.println("Resource #{++mSavedResourceCount} \"{args.ResourceFileName}\"");
            System.out.println("\t" + args.getResourceFileUri());

            // If we specified a ResourcesFolderAlias we will also need to redirect each stream to put its resource in that folder
            args.ResourceStream = new FileStream(args.getResourceFileUri(), FileMode.CREATE);
            args.setKeepResourceStreamOpen(false);
        }

        private int mSavedResourceCount;
    }
    //ExEnd
}
