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
import com.aspose.words.XamlFixedSaveOptions;
import org.testng.Assert;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.msConsole;
import com.aspose.words.IResourceSavingCallback;
import java.util.ArrayList;
import com.aspose.words.ResourceSavingArgs;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.msString;


@Test
public class ExXamlFixedSaveOptions extends ApiExampleBase
{
    //ExStart
    //ExFor:XamlFixedSaveOptions
    //ExFor:XamlFixedSaveOptions.ResourceSavingCallback
    //ExFor:XamlFixedSaveOptions.ResourcesFolder
    //ExFor:XamlFixedSaveOptions.ResourcesFolderAlias
    //ExFor:XamlFixedSaveOptions.SaveFormat
    //ExSummary:Shows how to print the URIs of linked resources created while converting a document to fixed-form .xaml.
    @Test //ExSkip
    public void resourceFolder() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        ResourceUriPrinter callback = new ResourceUriPrinter();

        // Create a "XamlFixedSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to the XAML save format.
        XamlFixedSaveOptions options = new XamlFixedSaveOptions();

        Assert.assertEquals(SaveFormat.XAML_FIXED, options.getSaveFormat());

        // Use the "ResourcesFolder" property to assign a folder in the local file system into which
        // Aspose.Words will save all the document's linked resources, such as images and fonts.
        options.setResourcesFolder(getArtifactsDir() + "XamlFixedResourceFolder");

        // Use the "ResourcesFolderAlias" property to use this folder
        // when constructing image URIs instead of the resources folder's name.
        options.setResourcesFolderAlias(getArtifactsDir() + "XamlFixedFolderAlias");

        options.setResourceSavingCallback(callback);

        // A folder specified by "ResourcesFolderAlias" will need to contain the resources instead of "ResourcesFolder".
        // We must ensure the folder exists before the callback's streams can put their resources into it.
        Directory.createDirectory(options.getResourcesFolderAlias());

        doc.save(getArtifactsDir() + "XamlFixedSaveOptions.ResourceFolder.xaml", options);

        for (String resource : callback.getResources())
            System.out.println(resource);
        testResourceFolder(callback); //ExSkip
    }

    /// <summary>
    /// Counts and prints URIs of resources created during conversion to fixed .xaml.
    /// </summary>
    private static class ResourceUriPrinter implements IResourceSavingCallback
    {
        public ResourceUriPrinter()
        {
            mResources = new ArrayList<String>();
        }

        public void /*IResourceSavingCallback.*/resourceSaving(ResourceSavingArgs args) throws Exception
        {
            getResources().add($"Resource \"{args.ResourceFileName}\"\n\t{args.ResourceFileUri}");

            // If we specified a resource folder alias, we would also need
            // to redirect each stream to put its resource in the alias folder.
            args.ResourceStream = new FileStream(args.getResourceFileUri(), FileMode.CREATE);
            args.setKeepResourceStreamOpen(false);
        }

        public ArrayList<String> getResources() { return mResources; };

        private ArrayList<String> mResources;
    }
    //ExEnd

    private void testResourceFolder(ResourceUriPrinter callback) throws Exception
    {
        Assert.assertEquals(15, callback.getResources().size());
        for (String resource : callback.getResources())
            Assert.assertTrue(File.exists(msString.split(resource, '\t')[1]));
    }
}
