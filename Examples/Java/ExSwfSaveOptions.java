//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.io.FileInputStream;


public class ExSwfSaveOptions extends ExBase
{
    @Test
    public void UseCustomToolTips() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        //ExStart
        //ExFor:SwfSaveOptions
        //ExFor:SwfSaveOptions.ToolTipsFontName
        //ExFor:SwfSaveOptions.ToolTips
        //ExFor:SwfViewerControlIdentifier
        //ExSummary:Shows how to change the the tooltips used in the embedded document viewer.
        // We create an instance of SwfSaveOptions to specify our custom tooltips.
        SwfSaveOptions options = new SwfSaveOptions();

        // By default, all tooltips are in English. You can specify font used for each tooltip.
        // Please note that font specified should contain proper glyphs normally used in tooltips.
        options.setToolTipsFontName("Times New Roman");

        // The following code will set the tooltip used for each control. In our case we will change the tooltips from English
        // to Russian.
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_ACTUAL_SIZE_BUTTON, "Оригинальный размер");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_FIT_TO_HEIGHT_BUTTON, "По высоте страницы");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_FIT_TO_WIDTH_BUTTON, "По ширине страницы");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_ZOOM_OUT_BUTTON, "Увеличить");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_ZOOM_OUT_BUTTON, "Уменшить");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_SELECTION_MODE_BUTTON, "Режим выделения текста");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_DRAG_MODE_BUTTON,"Режим перетаскивания");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_SINGLE_PAGE_SCROLL_LAYOUT_BUTTON, "Одностнаничный скролинг");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_SINGLE_PAGE_LAYOUT_BUTTON, "Одностраничный режим");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_TWO_PAGE_SCROLL_LAYOUT_BUTTON, "Двустраничный скролинг");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_TWO_PAGE_LAYOUT_BUTTON, "Двустраничный режим");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_TWO_PAGE_LAYOUT_BUTTON, "Полноэкранный режим");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_PREVIOUS_PAGE_BUTTON, "Предыдущая старница");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_PAGE_FIELD, "Введите номер страницы");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_NEXT_PAGE_BUTTON, "Следующая страница");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_SEARCH_FIELD, "Введите искомый текст");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.TOP_PANE_SEARCH_BUTTON, "Искать");

        // Left panel.
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.LEFT_PANE_DOCUMENT_MAP_BUTTON, "Карта документа");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.LEFT_PANE_PAGE_PREVIEW_PANE_BUTTON, "Предварительный просмотр страниц");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.LEFT_PANE_ABOUT_BUTTON, "О приложении");
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.LEFT_PANE_COLLAPSE_PANEL_BUTTON, "Свернуть панель");

        // Bottom panel.
        options.getToolTips().setBySwfViewerControlIdentifier(SwfViewerControlIdentifier.BOTTOM_PANE_SHOW_HIDE_BOTTOM_PANE_BUTTON, "Показать/Скрыть панель");
        //ExEnd

        doc.save(getMyDir() + "SwfSaveOptions.ToolTips Out.swf", options);
    }

    @Test
    public void HideControls() throws Exception
    {
        //ExStart
        //ExFor:SwfSaveOptions.TopPaneControlFlags
        //ExFor:SwfTopPaneControlFlags
        //ExFor:SwfSaveOptions.ShowSearch
        //ExSummary:Shows how to choose which controls to display in the embedded document viewer.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Create an instance of SwfSaveOptions and set some buttons as hidden.
        SwfSaveOptions options = new SwfSaveOptions();
        // Hide all the buttons with the exception of the page control buttons. Similar flags can be used for the left control pane as well.
        options.setTopPaneControlFlags(SwfTopPaneControlFlags.HIDE_ALL | SwfTopPaneControlFlags.SHOW_ACTUAL_SIZE |
                SwfTopPaneControlFlags.SHOW_FIT_TO_WIDTH | SwfTopPaneControlFlags.SHOW_FIT_TO_HEIGHT |
                SwfTopPaneControlFlags.SHOW_ZOOM_IN | SwfTopPaneControlFlags.SHOW_ZOOM_OUT);

        // You can also choose to show or hide the main elements of the viewer. Hide the search control.
        options.setShowSearch(false);
        //ExEnd

        doc.save(getMyDir() + "SwfSaveOptions.HideControls Out.swf", options);
    }

    @Test
    public void SetLogo() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        //ExStart
        //ExFor:SwfSaveOptions
        //ExFor:SwfSaveOptions.#ctor
        //ExFor:SwfSaveOptions.LogoImageBytes
        //ExFor:SwfSaveOptions.LogoLink
        //ExSummary:Shows how to specify a custom logo and link it to a web address in the embedded document viewer.
        // Create an instance of SwfSaveOptions.
        SwfSaveOptions options = new SwfSaveOptions();

        // Read the image from disk into byte array.
        FileInputStream fin = new FileInputStream(getMyDir() + "LogoSmall.png");
        byte[] logoBytes = new byte[fin.available()];
        fin.read(logoBytes);

        // Specify the logo image to use.
        options.setLogoImageBytes(logoBytes);

        // You can specify the URL of web page that should be opened when you click on the logo.
        options.setLogoLink("http://www.aspose.com");
        //ExEnd

        doc.save(getMyDir() + "SwfSaveOptions.CustomLogo Out.swf", options);
    }

}