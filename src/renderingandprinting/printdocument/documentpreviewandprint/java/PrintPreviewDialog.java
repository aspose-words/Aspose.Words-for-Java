/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package renderingandprinting.printdocument.documentpreviewandprint.java;

import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.PageRanges;
import javax.swing.*;
import javax.swing.plaf.basic.BasicArrowButton;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;
import java.awt.*;
import java.awt.event.*;
import java.awt.image.BufferedImage;
import java.awt.print.PageFormat;
import java.awt.print.Pageable;
import java.awt.print.Printable;
import java.awt.print.PrinterJob;


public class PrintPreviewDialog extends JFrame {
    /**
     * The dialog components.
     */
    private JButton printButton;
    private JComboBox zoomComboBox;
    private JProgressBar progressBar;
    private JButton closeButton;
    private JPanel contentPane;
    private JTextPane pageNumberTextBox;
    private BasicArrowButton previousPageButton;
    private BasicArrowButton firstPageButton;
    private BasicArrowButton lastPageButton;
    private BasicArrowButton nextPageButton;
    private JLabel imageLabel;
    private JScrollPane documentViewer;
    private JButton pageSetupButton;

    /**
     * Private instance members.
     */
    private Printable mPrintableDoc;
    private Pageable mPageableDoc;
    private boolean mIsOpened = true;
    private int mPreviousZoomIndex;
    private int mStartPage = 1;
    private int mTotalPages = -1;
    private int mCurrentPage = 1;
    private int mDocumentPages = -1;
    private boolean mPrintSelected = false;
    private PrinterJob mPrintJob;
    private PageFormat mPageFormat;
    private PrintRequestAttributeSet mAttributeSet;

    /**
     * Creates a new instance of PrintPreviewDialog for the given printable object. Since this object
     * does not define formatting for each page the preview dialog presents a page setup option
     * where the user can specify custom page settings for the document.
     * @param printJob The print job for the given document.
     * @param doc The printable document.
     */
    public PrintPreviewDialog(PrinterJob printJob, Printable doc)
    {
        mPrintableDoc = doc;
        mPrintJob = printJob;

        // In Java 1.6 and above we would use the getPageFormat(PrintRequestAttributeSet) property of a
        // class implementing printable to retrieve the currently specified page format. However this
        // property is not available in versions below versions so we need to assume the default page instead.
        mPageFormat = mPrintJob.defaultPage();

        init();
    }

    /**
     * Creates a new instance of PrintPreviewDialog for the given pageable object. Since this object
     * defines formatting for each page the page setup option on the preview dialog is disabled.
     * @param doc The pageable document.
     */
    public PrintPreviewDialog(Pageable doc)
    {
        mPageableDoc = doc;
        mTotalPages = doc.getNumberOfPages();
        mDocumentPages = mTotalPages;

        init();
    }

    public void setPrinterAttributes(PrintRequestAttributeSet attributes)
    {
        // Store the printer attributes for use with the page dialog.
        mAttributeSet = attributes;

        // Extract the page range from the printer attributes if that property is present.
        findPageRangeFromAttributes(attributes);
    }

    public void init()
    {
        // Setup the main window
        setContentPane(contentPane);
        pack();
        setTitle("Print preview");

        // Center the dialog in the center of the page.
        setLocationRelativeTo(null);

        // Add zoom options.
        populateZoomComboBox();

        // Pageable print classes already have page formatting applied so disable page setup button.
        if(mPageableDoc != null)
            pageSetupButton.setEnabled(false);

        // Make the page number centered horizontally.
        SimpleAttributeSet aSet = new SimpleAttributeSet();
        StyleConstants.setFontFamily(aSet, "Arial");
        StyleConstants.setAlignment(aSet, StyleConstants.ALIGN_CENTER);
        pageNumberTextBox.setParagraphAttributes(aSet, true);

        // Setup the appropriate handlers.
        printButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onPrint();
            }
        });

        pageSetupButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onPageSetup();
            }
        });

        zoomComboBox.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onZoomChanged();
            }
        });

        firstPageButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onFirstPageSelected();
            }
        });

        previousPageButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onPreviousPageSelected();
            }
        });

        lastPageButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onLastPageSelected();
            }
        });

        nextPageButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onNextPageSelected();
            }
        });

        closeButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                closeWindow();
            }
        });

        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                closeWindow();
            }
        });

        contentPane.registerKeyboardAction(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                closeWindow();
            }
        }, KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), JComponent.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT);
    }

    /**
     * Called to display the print preview dialog.
     */
    public boolean display()
    {
        // Activate the window.
        setVisible(true);

        // Render the current page.
        setPageToDisplay(mCurrentPage);

        // This causes the JFrame to act like modal. We want to use a JFrame in this way and not JDialog so
        // we can have a taskbar icon for this window.
        while(mIsOpened) {}

        // Return whether the user pressed print or close.
        return mPrintSelected;
    }

    /**
     * Renders the current page index of the document to image based on the current zoom factor and displays it on a JScrollPane.
     */
    private int renderImageAndDisplay()
    {
        // Set the progress bar to loading.
        progressBar.setVisible(true);
        progressBar.setIndeterminate(true);
        BufferedImage img = null;
        Graphics2D g;

        // The zoom factor currently selected.
        double zoomModifier = getCurrentZoomModifier();

        // Clear the current image.
        imageLabel.setIcon(null);

        int result;

        // Find the format of the current page from either the current pageable or printable object we are printing with.
        PageFormat format =  mPageableDoc != null ? mPageableDoc.getPageFormat(mCurrentPage - 1) : mPageFormat;

        try
        {
            img = new BufferedImage((int)(format.getWidth() * zoomModifier), (int)(format.getHeight() * zoomModifier), BufferedImage.TYPE_INT_RGB);
            g = img.createGraphics();

            // Fill the background white and add a black border.
            g.setColor(Color.WHITE);
            g.fillRect(0, 0, img.getWidth(), img.getHeight());
            g.setColor(Color.BLACK);
            g.drawRect(0, 0, img.getWidth() - 1, img.getHeight() - 1);

            // Scale based on zoom factor.
            g.scale(zoomModifier, zoomModifier);

            // We must re-size the image label so scrolling works properly.
            imageLabel.setPreferredSize(new Dimension(img.getWidth(), img.getHeight()));
            imageLabel.setMinimumSize(new Dimension(img.getWidth(), img.getHeight()));
            imageLabel.setMaximumSize(new Dimension(img.getWidth(), img.getHeight()));

            // Call the pageable or printable class to render the specified page onto our image object.
            if(mPageableDoc != null)
                result = mPageableDoc.getPrintable(mCurrentPage - 1).print(g, format, mCurrentPage - 1);
            else
                result = mPrintableDoc.print(g, format, mCurrentPage - 1);
        }

        catch(Exception e)
        {
            // We'll end up here if there is a problem with rendering or we have gone past the valid page range.
            // Display a blank page and return the result so we know we have gone past the last page.
            return Printable.NO_SUCH_PAGE;
        }

        finally
        {
            // Hide the progress bar.
            progressBar.setVisible(false);
        }

        // Display the rendered page.
        imageLabel.setIcon(new ImageIcon( img ));

        return result;
    }

    /**
     * Finds the page range selected by the user by the printer attributes.
     */
    private void findPageRangeFromAttributes(PrintRequestAttributeSet attributes)
    {
        if(attributes.containsKey(PageRanges.class))
        {
            int[][] pageRanges = ((PageRanges)attributes.get(PageRanges.class)).getMembers();

            int startPage = pageRanges[0][0];

            // If we know the number of pages in the document and the user specified value is out of range then use
            // the first page instead. Otherwise the user specified value is used which is checked later on if it's
            // valid or not.
            if(knowsDocumentPages())
            {
              if(startPage > mDocumentPages)
                mStartPage = 1;
              else
                mStartPage = startPage;
            }
            else
            {
                mStartPage = startPage;
            }
            
            mCurrentPage = mStartPage;

            // If we know how many pages the document has then we need to make sure the user user specified end page
            // does not go beyond this limit. Otherwise use this value which will be handled later on if it's found
            // to be invalid.
            if(knowsDocumentPages())
                mTotalPages = Math.min(mDocumentPages, pageRanges[0][1]);
            else
                mTotalPages = pageRanges[0][1];
        }
    }

    /**
     * Changes the current page to display.
     */
    private void setPageToDisplay(int page)
    {
        mCurrentPage = page;
        updateCurrentPage();
    }

    /**
     * Adds the zoom options to the combobox.
     */
    private void populateZoomComboBox()
    {
        zoomComboBox.addItem("10%");
        zoomComboBox.addItem("25%");
        zoomComboBox.addItem("50%");
        zoomComboBox.addItem("75%");
        zoomComboBox.addItem("100%");
        zoomComboBox.addItem("150%");
        zoomComboBox.addItem("200%");
        zoomComboBox.setSelectedIndex(4);

        mPreviousZoomIndex = zoomComboBox.getSelectedIndex();
    }

    /**
     * Returns the zoom modifier for the given zoom level.
     */
    private double getCurrentZoomModifier()
    {
        String zoomLevel = ((String)zoomComboBox.getSelectedItem()).replace("%", "");

        // Percent increase of the original image to make the page displayed at 100%.
        double hundredPercent = 1.5;

        // Zoom level percent as a decimal.
        double zoomValue = Double.parseDouble(zoomLevel) / 100;

        return hundredPercent * zoomValue;
    }

    /**
     * Verifies the page index based off the known valid page range and if valid calls for the new page to be rendered.
     * Depending if a page range was specified by the user we may or may not know the final page number yet. If we don't then
     * it is found later when each page is rendered.
     */
    private void updateCurrentPage()
    {
        // If we don't know the total pages in the document then we also won't know the final page count in the document
        // until we try to render outside the valid index. Until then allow the user to move forward no matter the current
        // page index.
        if(knowsDocumentPages())
        {
            if(mCurrentPage > mTotalPages)
                return;
        }

        // Can't move past the first page.
        if(mCurrentPage < mStartPage)
            return;

        int result = renderImageAndDisplay();

        // If the rendering of the previous page resulted in "NO_SUCH_PAGE" then
        // we must have found the last page in the document. The appropriate page numbers
        // needs to be updated.
        if(result == Printable.NO_SUCH_PAGE)
        {
            // If this occurs at the very start of the document preview then the intial starting page
            // must be out of bounds. Default to page one instead.
            if(mCurrentPage == mStartPage)
            {
                mCurrentPage = 1;
                mStartPage = 1;
            }
            else
            {
                mCurrentPage--;

                // If the next page document button was pressed and there's no futher pages to render then
                // we must have found the end page of the document.
                mTotalPages = mCurrentPage;
                mDocumentPages = mCurrentPage;
            }

            // Render the previous page which is the last page in the document.
            renderImageAndDisplay();
        }

        pageNumberTextBox.setText(String.valueOf((mCurrentPage - mStartPage) + 1));

        updateArrowButtons();
    }

    /**
     * Enables or disables arrow buttons based off the current state of the page index.
     */
    private void updateArrowButtons()
    {
        if(mCurrentPage > mStartPage)
        {
            firstPageButton.setEnabled(true);
            previousPageButton.setEnabled(true);
        }
        else
        {
            firstPageButton.setEnabled(false);
            previousPageButton.setEnabled(false);
        }

        if(knowsDocumentPages())
        {
            if(mCurrentPage == mTotalPages)
            {
                nextPageButton.setEnabled(false);
                lastPageButton.setEnabled(false);
            }
            else
            {
                nextPageButton.setEnabled(true);
                lastPageButton.setEnabled(true);
            }
        }
        else
        {
            // If we don't know how many pages are in the specified document then we cannot skip to the last page.
            lastPageButton.setEnabled(false);

            if(mCurrentPage == mTotalPages)
                nextPageButton.setEnabled(false);
            else
                nextPageButton.setEnabled(true);
        }
    }

   /**
     * Returns true if the last page number of the document is known.
     */
    private boolean knowsDocumentPages()
    {
        return mDocumentPages > 0;
    }

    /**
     * Called when user presses the cross on the window or the escape key to close the program.
     */
    private void closeWindow()
    {
        setVisible(false);
        mIsOpened = false;
        dispose();
    }

    /**
     * Called when the user press the Page Setup button. A screen is displayed which allows
     * the user to change the page setting of the document before printing.
     */
    private void onPageSetup()
    {
        // Retrieve the new page format from either the attributes if specified otherwise the previous page formatting.
        if(mAttributeSet != null)
           mPageFormat = mPrintJob.pageDialog(mAttributeSet);
        else
           mPageFormat = mPrintJob.pageDialog(mPageFormat);

        // Print using the new page format.
        mPrintJob.setPrintable(mPrintableDoc, mPageFormat);

        // Update the preview of the new settings.
        updateCurrentPage();
    }


    /**
     * Called when the user presses the print button. Returns true to specify that printing was accepted
     * and closes the window.
     */
    private void onPrint()
    {
        mPrintSelected = true;
        closeWindow();
    }

    /**
     * Re-renders the document page if the zoom level has been changed.
     */
    private void onZoomChanged()
    {
        if(mPreviousZoomIndex != zoomComboBox.getSelectedIndex())
        {
            mPreviousZoomIndex = zoomComboBox.getSelectedIndex();
            renderImageAndDisplay();
        }
    }

    /**
     * Called when the first page button is selected. Displays the first page.
     */
    private void onFirstPageSelected()
    {
        setPageToDisplay(mStartPage);
    }

    /**
     * Called when the next page button is selected. Displays the next page.
     */
    private void onNextPageSelected()
    {
        setPageToDisplay(mCurrentPage + 1);
    }

    /**
     * Called when the last page button is selected. Displays the last page.
     */
    private void onLastPageSelected()
    {
        setPageToDisplay(mTotalPages);
    }

    /**
     * Called when the previous page button is selected. Displays the previous page.
     */
    private void onPreviousPageSelected()
    {
        setPageToDisplay(mCurrentPage - 1);
    }

    /**
     * Creates the custom arrow buttons.
     */
    private void createUIComponents() {
        previousPageButton = new BasicArrowButton(BasicArrowButton.WEST);
        nextPageButton = new BasicArrowButton(BasicArrowButton.EAST);

        firstPageButton = new DoubleArrowButton(BasicArrowButton.WEST);
        lastPageButton = new DoubleArrowButton(BasicArrowButton.EAST);
    }

    /**
     * A simple extension of the BasicArrowButton which displays two arrows instead of one.
     */
    public class DoubleArrowButton extends BasicArrowButton
    {
        public DoubleArrowButton(int type) {
            super(type);
            mArrowType = type;
        }

        public void paintTriangle(Graphics g, int x, int y, int size,
                                  int direction, boolean isEnabled) {

            super.paintTriangle(g, x - (size / 2), y, size, mArrowType, isEnabled);
            super.paintTriangle(g, x + (size / 2), y, size, mArrowType, isEnabled);
        }

        private int mArrowType;
    }
}