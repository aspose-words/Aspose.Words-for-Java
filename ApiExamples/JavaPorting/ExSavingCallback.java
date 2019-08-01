package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.HtmlFixedSaveOptions;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.PsSaveOptions;
import com.aspose.words.SvgSaveOptions;
import com.aspose.words.XamlFixedSaveOptions;
import com.aspose.words.XpsSaveOptions;
import com.aspose.words.Document;
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.msString;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.IPageSavingCallback;
import com.aspose.words.PageSavingArgs;


@Test
class ExSavingCallback !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void checkThatAllMethodsArePresent()
    {
        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        htmlFixedSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        PsSaveOptions psSaveOptions = new PsSaveOptions();
        psSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        SvgSaveOptions svgSaveOptions = new SvgSaveOptions();
        svgSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        XamlFixedSaveOptions xamlFixedSaveOptions = new XamlFixedSaveOptions();
        xamlFixedSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        XpsSaveOptions xpsSaveOptions = new XpsSaveOptions();
        xpsSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());
    }

    @Test
    public void pageFileNameSavingCallback() throws Exception
    {
        //ExStart
        //ExFor:IPageSavingCallback
        //ExFor:PageSavingArgs
        //ExFor:PageSavingArgs.PageFileName
        //ExFor:FixedPageSaveOptions.PageSavingCallback
        //ExSummary:Shows how separate pages are saved when a document is exported to fixed page format.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        HtmlFixedSaveOptions htmlFixedSaveOptions =
            new HtmlFixedSaveOptions(); { htmlFixedSaveOptions.setPageIndex(0); htmlFixedSaveOptions.setPageCount(doc.getPageCount()); }
        htmlFixedSaveOptions.setPageSavingCallback(new CustomPageFileNamePageSavingCallback());

        doc.save(getArtifactsDir() + "Rendering.html", htmlFixedSaveOptions);

        String[] filePaths = Directory.getFiles(getArtifactsDir() + "", "Page_*.html");

        for (int i = 0; i < doc.getPageCount(); i++)
        {
            String file = msString.format(getArtifactsDir() + "Page_{0}.html", i);
            msAssert.areEqual(file, filePaths[i]);//ExSkip
        }
    }

    /// <summary>
    /// Custom PageFileName is specified.
    /// </summary>
    private static class CustomPageFileNamePageSavingCallback implements IPageSavingCallback
    {
        public void pageSaving(PageSavingArgs args)
        {
            // Specify name of the output file for the current page.
            args.setPageFileName(msString.format(getArtifactsDir() + "Page_{0}.html", args.getPageIndex()));
        }
    }
    //ExEnd
}
