package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.io.File;
import java.text.MessageFormat;

@Test
public class ExSvgSaveOptions extends ApiExampleBase {
    @Test
    public void saveLikeImage() throws Exception {
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
    public void svgResourceFolder() throws Exception {
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

        new File(options.getResourcesFolderAlias()).mkdir();

        doc.save(getArtifactsDir() + "SvgSaveOptions.SvgResourceFolder.svg", options);
    }

    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to .svg.
    /// </summary>
    private static class ResourceUriPrinter implements IResourceSavingCallback {
        public void resourceSaving(ResourceSavingArgs args) {
            System.out.println(MessageFormat.format("Resource #{0} \"{1}\"", ++mSavedResourceCount, args.getResourceFileName()));
            System.out.println("\t" + args.getResourceFileUri());
        }

        private int mSavedResourceCount;
    }
    //ExEnd

    @Test
    public void saveOfficeMath() throws Exception
    {
        //ExStart:SaveOfficeMath
        //GistId:9c17d666c47318436785490829a3984f
        //ExFor:NodeRendererBase.Save(String, SvgSaveOptions)
        //ExFor:NodeRendererBase.Save(Stream, SvgSaveOptions)
        //ExSummary:Shows how to pass save options when rendering office math.
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath math = (OfficeMath)doc.getChild(NodeType.OFFICE_MATH, 0, true);

        SvgSaveOptions options = new SvgSaveOptions();
        options.setTextOutputMode(SvgTextOutputMode.USE_PLACED_GLYPHS);

        math.getMathRenderer().save(getArtifactsDir() + "SvgSaveOptions.Output.svg", options);
        //ExEnd:SaveOfficeMath
    }

    @Test
    public void maxImageResolution() throws Exception
    {
        //ExStart:MaxImageResolution
        //GistId:f99d87e10ab87a581c52206321d8b617
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
        //GistId:c012c14781944ce4cc5e31f35b08060a
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
        //GistId:c012c14781944ce4cc5e31f35b08060a
        //ExFor:SvgSaveOptions.RemoveJavaScriptFromLinks
        //ExSummary:Shows how to remove JavaScript from the links (svg).
        Document doc = new Document(getMyDir() + "JavaScript in HREF.docx");

        SvgSaveOptions saveOptions = new SvgSaveOptions();
        saveOptions.setRemoveJavaScriptFromLinks(true);

        doc.save(getArtifactsDir() + "SvgSaveOptions.RemoveJavaScriptFromLinksSvg.html", saveOptions);
        //ExEnd:RemoveJavaScriptFromLinksSvg
    }
}
