package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.text.MessageFormat;
import java.util.ArrayList;

@Test
public class ExXamlFixedSaveOptions extends ApiExampleBase {
    //ExStart
    //ExFor:XamlFixedSaveOptions
    //ExFor:XamlFixedSaveOptions.ResourceSavingCallback
    //ExFor:XamlFixedSaveOptions.ResourcesFolder
    //ExFor:XamlFixedSaveOptions.ResourcesFolderAlias
    //ExFor:XamlFixedSaveOptions.SaveFormat
    //ExSummary:Shows how to print the URIs of linked resources created during conversion of a document to fixed-form .xaml.
    @Test //ExSkip
    public void resourceFolder() throws Exception {
        // Open a document which contains resources
        Document doc = new Document(getMyDir() + "Rendering.docx");

        ResourceUriPrinter callback = new ResourceUriPrinter();

        XamlFixedSaveOptions options = new XamlFixedSaveOptions();
        {
            options.setSaveFormat(SaveFormat.XAML_FIXED);
            options.setResourcesFolder(getArtifactsDir() + "XamlFixedResourceFolder");
            options.setResourcesFolderAlias(getArtifactsDir() + "XamlFixedFolderAlias");
            options.setResourceSavingCallback(callback);
        }

        // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder
        // We must ensure the folder exists before the streams can put their resources into it
        new File(options.getResourcesFolderAlias()).mkdir();

        doc.save(getArtifactsDir() + "XamlFixedSaveOptions.ResourceFolder.xaml", options);

        for (String resource : callback.getResources())
            System.out.println(resource);
    }

    /// <summary>
    /// Counts and prints URIs of resources created during conversion to to fixed .xaml.
    /// </summary>
    private static class ResourceUriPrinter implements IResourceSavingCallback {
        public ResourceUriPrinter() {
            mResources = new ArrayList<String>();
        }

        public void resourceSaving(ResourceSavingArgs args) throws Exception {
            // If we set a folder alias in the SaveOptions object, it will be stored here
            getResources().add(MessageFormat.format("Resource \"{0}\"\n\t{1}", args.getResourceFileName(), args.getResourceFileUri()));

            // If we specified a ResourcesFolderAlias we will also need to redirect each stream to put its resource in that folder
            args.setResourceStream(new FileOutputStream(args.getResourceFileUri()));
            args.setKeepResourceStreamOpen(false);
        }

        public ArrayList<String> getResources() {
            return mResources;
        }

        ;

        private ArrayList<String> mResources;
    }
    //ExEnd
}
