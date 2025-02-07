// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.OfficeMath;
import com.aspose.words.NodeType;
import com.aspose.ms.System.IO.MemoryStream;


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

    @Test
    public void saveOfficeMath() throws Exception
    {
        //ExStart:SaveOfficeMath
        //GistId:a775441ecb396eea917a2717cb9e8f8f
        //ExFor:NodeRendererBase.Save(String, SvgSaveOptions)
        //ExFor:NodeRendererBase.Save(Stream, SvgSaveOptions)
        //ExSummary:Shows how to pass save options when rendering office math.
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath math = (OfficeMath)doc.getChild(NodeType.OFFICE_MATH, 0, true);

        SvgSaveOptions options = new SvgSaveOptions();
        options.setTextOutputMode(SvgTextOutputMode.USE_PLACED_GLYPHS);

        math.getMathRenderer().save(getArtifactsDir() + "SvgSaveOptions.Output.svg", options);
        
        MemoryStream stream = new MemoryStream();
        try /*JAVA: was using*/
    	{
            math.getMathRenderer().save(stream, options);
    	}
        finally { if (stream != null) stream.close(); }
        //ExEnd:SaveOfficeMath
    }

    @Test
    public void maxImageResolution() throws Exception
    {
        //ExStart:MaxImageResolution
        //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
        //ExFor:ShapeBase.SoftEdge
        //ExFor:SoftEdgeFormat.Radius
        //ExFor:SoftEdgeFormat.Remove
        //ExFor:SvgSaveOptions.MaxImageResolution
        //ExSummary:Shows how to set limit for image resolution.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        SvgSaveOptions saveOptions = new SvgSaveOptions();
        saveOptions.setMaxImageResolution(72);

        doc.save(getArtifactsDir() + "SvgSaveOptions.MaxImageResolution.svg", saveOptions);
        //ExEnd:MaxImageResolution
    }

    @Test
    public void idPrefixSvg() throws Exception
    {
        //ExStart:IdPrefixSvg
        //GistId:f86d49dc0e6781b93e576539a01e6ca2
        //ExFor:SvgSaveOptions.IdPrefix
        //ExSummary:Shows how to add a prefix that is prepended to all generated element IDs (svg).
        Document doc = new Document(getMyDir() + "Id prefix.docx");

        SvgSaveOptions saveOptions = new SvgSaveOptions();
        saveOptions.setIdPrefix("pfx1_");

        doc.save(getArtifactsDir() + "SvgSaveOptions.IdPrefixSvg.html", saveOptions);
        //ExEnd:IdPrefixSvg
    }

    @Test
    public void removeJavaScriptFromLinksSvg() throws Exception
    {
        //ExStart:RemoveJavaScriptFromLinksSvg
        //GistId:f86d49dc0e6781b93e576539a01e6ca2
        //ExFor:SvgSaveOptions.RemoveJavaScriptFromLinks
        //ExSummary:Shows how to remove JavaScript from the links (svg).
        Document doc = new Document(getMyDir() + "JavaScript in HREF.docx");

        SvgSaveOptions saveOptions = new SvgSaveOptions();
        saveOptions.setRemoveJavaScriptFromLinks(true);

        doc.save(getArtifactsDir() + "SvgSaveOptions.RemoveJavaScriptFromLinksSvg.html", saveOptions);
        //ExEnd:RemoveJavaScriptFromLinksSvg
    }
}

