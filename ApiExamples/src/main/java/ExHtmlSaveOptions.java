//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.annotations.DataProvider;

import java.io.File;
import java.util.ArrayList;
import java.util.regex.Pattern;

public class ExHtmlSaveOptions extends ApiExampleBase
{
    //For assert this test you need to open HTML docs and they shouldn't have negative left margins
    @Test(dataProvider = "exportPageMarginsDataProvider")
    public void exportPageMargins(/*SaveFormat*/int saveFormat) throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setSaveFormat(saveFormat);
        saveOptions.setExportPageMargins(true);

        save(doc, "\\Artifacts\\HtmlSaveOptions.ExportPageMargins." + SaveFormat.toString(saveFormat).toLowerCase(), saveFormat, saveOptions);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportPageMarginsDataProvider")
    public static Object[][] exportPageMarginsDataProvider() throws Exception
    {
        return new Object[][]{{SaveFormat.HTML}, {SaveFormat.MHTML}, {SaveFormat.EPUB},};
    }

    @Test(dataProvider = "exportOfficeMathDataProvider")
    public void exportOfficeMath(/*SaveFormat*/int saveFormat, /*HtmlOfficeMathOutputMode*/int outputMode) throws Exception
    {
        Document doc = new Document(getMyDir() + "OfficeMath.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setOfficeMathOutputMode(outputMode);

        save(doc, "\\Artifacts\\HtmlSaveOptions.ExportToHtmlUsingImage." + SaveFormat.toString(saveFormat).toLowerCase(), saveFormat, saveOptions);

        switch (saveFormat)
        {
            case SaveFormat.HTML:
                DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportToHtmlUsingImage." + SaveFormat.toString(saveFormat).toLowerCase(), "<img src=\"HtmlSaveOptions.ExportToHtmlUsingImage.001.png\" width=\"49\" height=\"19\" alt=\"\" style=\"vertical-align:middle; -aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />");
                return;

            case SaveFormat.MHTML:
                DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportToHtmlUsingImage." + SaveFormat.toString(saveFormat).toLowerCase(), "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>A</mi><mo>=</mo><mi>π</mi><msup><mrow><mi>r</mi></mrow><mrow><mn>2</mn></mrow></msup></math>");
                return;

            case SaveFormat.EPUB:
                DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportToHtmlUsingImage." + SaveFormat.toString(saveFormat).toLowerCase(), "<span style=\"font-family:'Cambria Math'\">A=π</span><span style=\"font-family:'Cambria Math'\">r</span><span style=\"font-family:'Cambria Math'\">2</span>");
                return;
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportOfficeMathDataProvider")
    public static Object[][] exportOfficeMathDataProvider() throws Exception
    {
        return new Object[][]{{SaveFormat.HTML, HtmlOfficeMathOutputMode.IMAGE}, {SaveFormat.MHTML, HtmlOfficeMathOutputMode.MATH_ML}, {SaveFormat.EPUB, HtmlOfficeMathOutputMode.TEXT},};
    }

    @Test(dataProvider = "exportTextBoxAsSvgDataProvider")
    public void exportTextBoxAsSvg(/*SaveFormat*/int saveFormat, boolean textBoxAsSvg) throws Exception
    {
        ArrayList<String> dirFiles;

        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportTextBoxAsSvg.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportTextBoxAsSvg(textBoxAsSvg);

        save(doc, "\\Artifacts\\HtmlSaveOptions.ExportTextBoxAsSvg." + SaveFormat.toString(saveFormat).toLowerCase(), saveFormat, saveOptions);

        switch (saveFormat)
        {
            case SaveFormat.HTML:

                dirFiles = DirectoryGetFiles(getMyDir() + "\\Artifacts\\", "HtmlSaveOptions.ExportTextBoxAsSvg.001.png");
                Assert.assertTrue(dirFiles.isEmpty());

                DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportTextBoxAsSvg." + SaveFormat.toString(saveFormat).toLowerCase(), "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"238\" height=\"185\"><defs><clipPath id=\"clip1\"><path d=\"M0,4.800000191 L238.293334961,4.800000191 L238.293334961,113.006721497 L0,113.006721497 Z\" clip-rule=\"evenodd\" /></clipPath></defs><g><g><g transform=\"matrix(1,0,0,1,0,0)\"><path d=\"M0,0 L238.293334961,0 L238.293334961,0 L238.293334961,117.806724548 L238.293334961,117.806724548 L0,117.806724548 Z\" fill=\"#ffffff\" fill-rule=\"evenodd\" /><path d=\"M0,0 L238.293334961,0 L238.293334961,0 L238.293334961,117.806724548 L238.293334961,117.806724548 L0,117.806724548 Z\" stroke-width=\"1\" stroke-miterlimit=\"13.333333015\" stroke=\"#000000\" fill=\"none\" fill-rule=\"evenodd\" /><g transform=\"matrix(1,0,0,1,0,0)\" clip-path=\"url(#clip1)\"><g transform=\"matrix(1,0,0,1,10.066666603,5.266666889)\"><text><tspan x=\"0\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">[Grab</tspan><tspan x=\"33.594665527\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"36.910667419\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">your</tspan><tspan x=\"64.102668762\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"67.418670654\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">reader’s</tspan><tspan x=\"116.366668701\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"119.682670593\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">attention</tspan><tspan x=\"175.255996704\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"178.572006226\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">with</tspan><tspan x=\"205.04133606\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"208.357345581\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">a</tspan><tspan x=\"215.382675171\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">great</tspan><tspan x=\"31.251998901\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"34.568004608\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">quote</tspan><tspan x=\"69.924003601\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"73.239997864\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">from</tspan><tspan x=\"102.279998779\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"105.596000671\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">the</tspan><tspan x=\"125.512016296\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"128.82800293\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">document</tspan><tspan x=\"189.807998657\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"193.124008179\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">or</tspan><tspan x=\"205.972000122\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">use</tspan><tspan x=\"20.739999771\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"24.055999756\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">this</tspan><tspan x=\"45.777332306\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"49.093334198\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">space</tspan><tspan x=\"83.060005188\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"86.375999451\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">to</tspan><tspan x=\"99.022666931\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"102.338661194\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">emphasize</tspan><tspan x=\"165.982666016\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"169.298660278\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">a</tspan><tspan x=\"176.323989868\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"179.63999939\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">key</tspan><tspan x=\"200.244003296\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">point.</tspan><tspan x=\"35.126667023\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"38.442668915\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">To</tspan><tspan x=\"53.324001312\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"56.63999939\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">place</tspan><tspan x=\"88.236000061\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"91.552001953\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">this</tspan><tspan x=\"113.273338318\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"116.589332581\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">text</tspan><tspan x=\"140.063995361\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"143.380004883\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">box</tspan><tspan x=\"165.17199707\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">anywhere</tspan><tspan x=\"59.268001556\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"62.584003448\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">on</tspan><tspan x=\"78.024002075\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"81.340003967\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">the</tspan><tspan x=\"101.256004333\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"104.571998596\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">page,</tspan><tspan x=\"137.164001465\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"140.479995728\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">just</tspan><tspan x=\"162.344009399\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"165.660003662\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">drag</tspan><tspan x=\"192.408004761\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"195.723999023\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">it.]</tspan></text></g></g></g></g></g></svg>");
                return;

            case SaveFormat.EPUB:

                dirFiles = DirectoryGetFiles(getMyDir() + "\\Artifacts\\", "HtmlSaveOptions.ExportTextBoxAsSvg.001.png");
                Assert.assertTrue(dirFiles.isEmpty());

                DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportTextBoxAsSvg." + SaveFormat.toString(saveFormat).toLowerCase(), "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"238\" height=\"185\"><defs><clipPath id=\"clip1\"><path d=\"M0,4.800000191 L238.293334961,4.800000191 L238.293334961,113.006721497 L0,113.006721497 Z\" clip-rule=\"evenodd\" /></clipPath></defs><g><g><g transform=\"matrix(1,0,0,1,0,0)\"><path d=\"M0,0 L238.293334961,0 L238.293334961,0 L238.293334961,117.806724548 L238.293334961,117.806724548 L0,117.806724548 Z\" fill=\"#ffffff\" fill-rule=\"evenodd\" /><path d=\"M0,0 L238.293334961,0 L238.293334961,0 L238.293334961,117.806724548 L238.293334961,117.806724548 L0,117.806724548 Z\" stroke-width=\"1\" stroke-miterlimit=\"13.333333015\" stroke=\"#000000\" fill=\"none\" fill-rule=\"evenodd\" /><g transform=\"matrix(1,0,0,1,0,0)\" clip-path=\"url(#clip1)\"><g transform=\"matrix(1,0,0,1,10.066666603,5.266666889)\"><text><tspan x=\"0\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">[Grab</tspan><tspan x=\"33.594665527\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"36.910667419\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">your</tspan><tspan x=\"64.102668762\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"67.418670654\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">reader’s</tspan><tspan x=\"116.366668701\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"119.682670593\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">attention</tspan><tspan x=\"175.255996704\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"178.572006226\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">with</tspan><tspan x=\"205.04133606\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"208.357345581\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">a</tspan><tspan x=\"215.382675171\" y=\"13.965332985\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">great</tspan><tspan x=\"31.251998901\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"34.568004608\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">quote</tspan><tspan x=\"69.924003601\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"73.239997864\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">from</tspan><tspan x=\"102.279998779\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"105.596000671\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">the</tspan><tspan x=\"125.512016296\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"128.82800293\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">document</tspan><tspan x=\"189.807998657\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"193.124008179\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">or</tspan><tspan x=\"205.972000122\" y=\"33.28666687\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">use</tspan><tspan x=\"20.739999771\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"24.055999756\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">this</tspan><tspan x=\"45.777332306\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"49.093334198\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">space</tspan><tspan x=\"83.060005188\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"86.375999451\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">to</tspan><tspan x=\"99.022666931\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"102.338661194\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">emphasize</tspan><tspan x=\"165.982666016\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"169.298660278\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">a</tspan><tspan x=\"176.323989868\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"179.63999939\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">key</tspan><tspan x=\"200.244003296\" y=\"52.608001709\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">point.</tspan><tspan x=\"35.126667023\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"38.442668915\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">To</tspan><tspan x=\"53.324001312\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"56.63999939\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">place</tspan><tspan x=\"88.236000061\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"91.552001953\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">this</tspan><tspan x=\"113.273338318\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"116.589332581\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">text</tspan><tspan x=\"140.063995361\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"143.380004883\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">box</tspan><tspan x=\"165.17199707\" y=\"71.929328918\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"0\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">anywhere</tspan><tspan x=\"59.268001556\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"62.584003448\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">on</tspan><tspan x=\"78.024002075\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"81.340003967\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">the</tspan><tspan x=\"101.256004333\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"104.571998596\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">page,</tspan><tspan x=\"137.164001465\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"140.479995728\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">just</tspan><tspan x=\"162.344009399\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"165.660003662\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">drag</tspan><tspan x=\"192.408004761\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\"> </tspan><tspan x=\"195.723999023\" y=\"91.250671387\" font-family=\"Calibri\" font-weight=\"normal\" font-style=\"normal\" font-size=\"14.666666985\" fill=\"#000000\">it.]</tspan></text></g></g></g></g></g></svg>");
                return;

            case SaveFormat.MHTML:

                dirFiles = DirectoryGetFiles(getMyDir() + "\\Artifacts\\", "HtmlSaveOptions.ExportTextBoxAsSvg.001.png");
                Assert.assertFalse(dirFiles.isEmpty());

                DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportTextBoxAsSvg." + SaveFormat.toString(saveFormat).toLowerCase(), "<img src=\"HtmlSaveOptions.ExportTextBoxAsSvg.001.png\" width=\"241\" height=\"120\" alt=\"\" style=\"margin:3.6pt 9pt; -aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:14.4pt; -aw-wrap-type:square; float:left\" />");
                return;
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportTextBoxAsSvgDataProvider")
    public static Object[][] exportTextBoxAsSvgDataProvider() throws Exception
    {
        return new Object[][]{{SaveFormat.HTML, true}, {SaveFormat.EPUB, true}, {SaveFormat.MHTML, false},};
    }

    private ArrayList<String> DirectoryGetFiles(String dirname, String filenamePattern)
    {
        File dirFile = new File(dirname);
        Pattern re = Pattern.compile(filenamePattern.replace("*", ".*").replace("?", ".?"));
        ArrayList<String> dirFiles = new ArrayList<String>();
        for (File file : dirFile.listFiles())
        {
            if (file.isDirectory()) dirFiles.addAll(DirectoryGetFiles(file.getPath(), filenamePattern));
            else if (re.matcher(file.getName()).matches()) dirFiles.add(file.getPath());
        }
        return dirFiles;
    }

    private static Document save(Document inputDoc, String outputDocPath, /*SaveFormat*/int saveFormat, SaveOptions saveOptions) throws Exception
    {
        switch (saveFormat)
        {
            case SaveFormat.HTML:
                inputDoc.save(getMyDir() + outputDocPath, saveOptions);
                return inputDoc;
            case SaveFormat.MHTML:
                inputDoc.save(getMyDir() + outputDocPath, saveOptions);
                return inputDoc;
            case SaveFormat.EPUB:
                inputDoc.save(getMyDir() + outputDocPath, saveOptions);
                return inputDoc;
        }

        return inputDoc;
    }

    @Test
    public void controlListLabelsExportToHtml() throws Exception
    {
        Document doc = new Document(getMyDir() + "Lists.PrintOutAllLists.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);

        // This option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss, 
        // otherwise HTML <p> tag is used. This is also the default value.
        saveOptions.setExportListLabels(ExportListLabels.AUTO);
        doc.save(getMyDir() + "\\Artifacts\\Document.ExportListLabels Auto.html", saveOptions);

        // Using this option the <p> tag is used for any list label representation.
        saveOptions.setExportListLabels(ExportListLabels.AS_INLINE_TEXT);
        doc.save(getMyDir() + "\\Artifacts\\Document.ExportListLabels InlineText.html", saveOptions);

        // The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
        saveOptions.setExportListLabels(ExportListLabels.BY_HTML_TAGS);
        doc.save(getMyDir() + "\\Artifacts\\Document.ExportListLabels HtmlTags.html", saveOptions);
    }

    @Test(dataProvider = "exportUrlForLinkedImageDataProvider")
    public void exportUrlForLinkedImage(boolean export) throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportOriginalUrlForLinkedImages(export);

        doc.save(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

        ArrayList<String> dirFiles = DirectoryGetFiles(getMyDir() + "\\Artifacts\\", "HtmlSaveOptions.ExportUrlForLinkedImage.001.png");

        if (dirFiles.size() == 0)
            DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
        else
            DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportUrlForLinkedImageDataProvider")
    public static Object[][] exportUrlForLinkedImageDataProvider() throws Exception
    {
        return new Object[][]{{true}, {false},};
    }

    @Test(enabled = false, description = "Bug, css styles starting with -aw, even if ExportRoundtripInformation is false", dataProvider = "exportRoundtripInformationDataProvider")
    public void exportRoundtripInformation(boolean valueHtml) throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportRoundtripInformation(valueHtml);

        doc.save(getMyDir() + "\\Artifacts\\HtmlSaveOptions.RoundtripInformation.html");

        if (valueHtml)
            DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.RoundtripInformation.html", "<img src=\"HtmlSaveOptions.RoundtripInformation.003.png\" width=\"226\" height=\"132\" alt=\"\" style=\"margin-top:-53.74pt; margin-left:-26.75pt; -aw-left-pos:-26.25pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:41.25pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:1\"><img src=\"HtmlSaveOptions.RoundtripInformation.002.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:74.51pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:169.5pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:2\"><img src=\"HtmlSaveOptions.RoundtripInformation.001.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:199.01pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:294pt; -aw-wrap-type:none; position:absolute\" />");
        else
            DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.RoundtripInformation.html", "<img src=\"HtmlSaveOptions.RoundtripInformation.003.png\" width=\"226\" height=\"132\" alt=\"\" style=\"margin-top:-53.74pt; margin-left:-26.75pt; -aw-left-pos:-26.25pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:41.25pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:1\"><img src=\"HtmlSaveOptions.RoundtripInformation.002.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:74.51pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:169.5pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:2\"><img src=\"HtmlSaveOptions.RoundtripInformation.001.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:199.01pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:294pt; -aw-wrap-type:none; position:absolute\" />");
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportRoundtripInformationDataProvider")
    public static Object[][] exportRoundtripInformationDataProvider() throws Exception
    {
        return new Object[][]{{true}, {false},};
    }

    @Test
    public void roundtripInformationDefaulValue()
    {
        //Assert that default value is true for HTML and false for MHTML and EPUB.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), true);

        saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), false);

        saveOptions = new HtmlSaveOptions(SaveFormat.EPUB);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), false);
    }

    @Test
    public void configForSavingExternalResources() throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setResourceFolder("Resources");
        saveOptions.setResourceFolderAlias("https://www.aspose.com/");

        doc.save(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportPageMargins Out.html", saveOptions);

        ArrayList<String> imageFiles = DirectoryGetFiles(getMyDir() + "\\Artifacts\\Resources\\", "*.png");
        Assert.assertEquals(3, imageFiles.size());

        ArrayList<String> fontFiles = DirectoryGetFiles(getMyDir() + "\\Artifacts\\Resources\\", "*.ttf");
        Assert.assertEquals(1, fontFiles.size());

        ArrayList<String> cssFiles = DirectoryGetFiles(getMyDir() + "\\Artifacts\\Resources\\", "*.css");
        Assert.assertEquals(1, cssFiles.size());

        DocumentHelper.findTextInFile(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportPageMargins Out.html", "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExportPageMargins Out.css\"");
    }

    @Test
    public void convertFontsAsBase64() throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setResourceFolder("Resources");
        saveOptions.setExportFontResources(true);
        saveOptions.setExportFontsAsBase64(true);

        doc.save(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ExportPageMargins Out.html", saveOptions);
    }

    @Test(dataProvider = "html5SupportDataProvider")
    public void html5Support(/*HtmlVersion*/int htmlVersion) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setHtmlVersion(htmlVersion);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "html5SupportDataProvider")
    public static Object[][] html5SupportDataProvider() throws Exception
    {
        return new Object[][]{{HtmlVersion.HTML_5}, {HtmlVersion.XHTML},};
    }

    @Test(dataProvider = "exportFontsDataProvider")
    public void exportFonts(boolean exportAsBase64) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportFontResources(true);
        saveOptions.setExportFontsAsBase64(exportAsBase64);

        if (!exportAsBase64)
        {
            doc.save(getMyDir() + "\\Artifacts\\DocumentExportFonts Out 1.html", saveOptions);
            Assert.assertFalse(DirectoryGetFiles(getMyDir() + "\\Artifacts\\", "DocumentExportFonts Out 1.times.ttf").isEmpty()); //Verify that the font has been added to the folder

        } else
        {
            doc.save(getMyDir() + "\\Artifacts\\DocumentExportFonts Out 2.html", saveOptions);
            Assert.assertTrue(DirectoryGetFiles(getMyDir() + "\\Artifacts\\", "DocumentExportFonts Out 2.times.ttf").isEmpty()); //Verify that the font is not added to the folder

        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportFontsDataProvider")
    public static Object[][] exportFontsDataProvider() throws Exception
    {
        return new Object[][]{{false}, {true},};
    }

    @Test
    public void resourceFolderPriority() throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ResourceFolder.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setResourceFolder(getMyDir() + "\\Artifacts\\Resources");
        saveOptions.setResourceFolderAlias("http://example.com/resources");

        doc.save(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ResourceFolder Out.html", saveOptions);

        Assert.assertFalse(DirectoryGetFiles(getMyDir() + "\\Artifacts\\Resources", "HtmlSaveOptions.ResourceFolder Out.001.jpeg").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getMyDir() + "\\Artifacts\\Resources", "HtmlSaveOptions.ResourceFolder Out.002.png").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getMyDir() + "\\Artifacts\\Resources", "HtmlSaveOptions.ResourceFolder Out.calibri.ttf").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getMyDir() + "\\Artifacts\\Resources", "HtmlSaveOptions.ResourceFolder Out.css").isEmpty());

    }

    @Test
    public void resourceFolderLowPriority() throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ResourceFolder.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setFontsFolder(getMyDir() + "\\Artifacts\\Fonts");
        saveOptions.setImagesFolder(getMyDir() + "\\Artifacts\\Images");
        saveOptions.setResourceFolder(getMyDir() + "\\Artifacts\\Resources");
        saveOptions.setResourceFolderAlias("http://example.com/resources");

        doc.save(getMyDir() + "\\Artifacts\\HtmlSaveOptions.ResourceFolder Out.html", saveOptions);

        Assert.assertFalse(DirectoryGetFiles(getMyDir() + "\\Artifacts\\Images", "HtmlSaveOptions.ResourceFolder Out.001.jpeg").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getMyDir() + "\\Artifacts\\Images", "HtmlSaveOptions.ResourceFolder Out.002.png").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getMyDir() + "\\Artifacts\\Fonts", "HtmlSaveOptions.ResourceFolder Out.calibri.ttf").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getMyDir() + "\\Artifacts\\Resources", "HtmlSaveOptions.ResourceFolder Out.css").isEmpty());
    }
}
