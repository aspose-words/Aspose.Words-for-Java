/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package viewersandvisualizers.documentexplorer.java;

import com.aspose.words.Document;

import javax.swing.*;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;

/**
* This class is used as a repository for objects, that should be available from any place of the project code.
*/
public class Globals
{
    /**
    * This class is purely static, that's why we prevent instance creation by declaring the constructor as private.
    */
    private Globals() {}

	// Titles used within the application.
	static final String APPLICATION_TITLE = "Document Explorer";
	static final String UNEXPECTED_EXCEPTION_DIALOG_TITLE = APPLICATION_TITLE + " - unexpected error occured";
	static final String OPEN_DOCUMENT_DIALOG_TITLE = "Open Document";
	static final String SAVE_DOCUMENT_DIALOG_TITLE = "Save Document As";

    // Open File filters
	static final OpenFileFilter OPEN_FILE_FILTER_ALL_SUPPORTED_FORMATS = new OpenFileFilter(
			new String[] {".doc",".dot",".docx",".dotx",".docm",".dotm",".xml",".wml",".rtf",".odt",".ott",".htm",".html",".xhtml",".mht",".mhtm",".mhtml"}, "All Supported Formats (*.doc;*.dot;*.docx;*.dotx;*.docm;*.dotm;*.xml;*.wml;*.rtf;*.odt;*.ott;*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml)");

    static final OpenFileFilter OPEN_FILE_FILTER_DOC_FORMAT = new OpenFileFilter(
			new String[] {".doc", ".doct"}, "Word 97-2003 Documents (*.doc;*.dot)");

    static final OpenFileFilter OPEN_FILE_FILTER_DOCX_FORMAT = new OpenFileFilter(
			new String[] {".docx", ".dotx", ".docm", ".dotm"}, "Word 2007 OOXML Documents (*.docx;*.dotx;*.docm;*.dotm)");

    static final OpenFileFilter OPEN_FILE_FILTER_XML_FORMAT = new OpenFileFilter(
			new String[] {".xml", ".wml"}, "XML Documents (*.xml;*.wml)");

    static final OpenFileFilter OPEN_FILE_FILTER_RTF_FORMAT = new OpenFileFilter(
			new String[] {".rtf"}, "Rich Text Format (*.rtf)");

    static final OpenFileFilter OPEN_FILE_FILTER_ODT_FORMAT = new OpenFileFilter(
			new String[] {".odt", ".ott"}, "OpenDocument Text (*.odt;*.ott)");

    static final OpenFileFilter OPEN_FILE_FILTER_HTML_FORMAT = new OpenFileFilter(
			new String[] {".htm", ".html", ".xhtml", ".mht", ".mhtm", ".mhtml"}, "Web Pages (*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml)");

    // Save File Filters
    static final SaveFileFilter SAVE_FILE_FILTER_DOC = new SaveFileFilter(
            ".doc", "Word 97-2003 Document (*.doc)");

    static final SaveFileFilter SAVE_FILE_FILTER_DOCX = new SaveFileFilter(
            ".docx", "Word 2007 OOXML Document (*.docx)");

    static final SaveFileFilter SAVE_FILE_FILTER_DOCM = new SaveFileFilter(
            ".docm", "Word 2007 OOXML Macro-Enabled Document (*.docm)");

    static final SaveFileFilter SAVE_FILE_FILTER_PDF = new SaveFileFilter(
            ".pdf", "PDF (*.pdf)");

    static final SaveFileFilter SAVE_FILE_FILTER_XPS = new SaveFileFilter(
            ".xps", "XPS Document (*.xps)");

    static final SaveFileFilter SAVE_FILE_FILTER_PDT = new SaveFileFilter(
            ".odt", "OpenDocument Text (*.odt)");

    static final SaveFileFilter SAVE_FILE_FILTER_HTML = new SaveFileFilter(
            ".html", "Web Page (*.html)");

    static final SaveFileFilter SAVE_FILE_FILTER_MHT = new SaveFileFilter(
            ".mht", "Single File Web Page (*.mht)");

    static final SaveFileFilter SAVE_FILE_FILTER_RTF = new SaveFileFilter(
            ".rtf", "Rich Text Format (*.rtf)");

    static final SaveFileFilter SAVE_FILE_FILTER_XML = new SaveFileFilter(
            ".xml", "Word 2003 WordprocessingML (*.xml)");

    static final SaveFileFilter SAVE_FILE_FILTER_FOPC = new SaveFileFilter(
            ".fopc", "FlatOPC XML Document (*.fopc)");

    static final SaveFileFilter SAVE_FILE_FILTER_TXT = new SaveFileFilter(
            ".txt", "Plain Text (*.txt)");

    static final SaveFileFilter SAVE_FILE_FILTER_EPUB = new SaveFileFilter(
            ".epub", "IDPF EPUB Document (*.epub)");

    static final SaveFileFilter SAVE_FILE_FILTER_SWF = new SaveFileFilter(
            ".swf", "Macromedia Flash File (*.swf)");

    static final SaveFileFilter SAVE_FILE_FILTER_XAML = new SaveFileFilter(
            ".xaml", "XAML Fixed Document (*.xaml)");

	/**
	* Reference for application's main form.
	*/
	static MainForm mMainForm;

    /**
    * Reference for currently loaded Document.
    */
    static Document mDocument;

    /**
    * Reference for current Tree Model
    */
    static DefaultTreeModel mTreeModel;

    /**
    * Reference for the current Tree
    */
    static JTree mTree;

    /**
    * Reference for the current root node.
    */
    static DefaultMutableTreeNode mRootNode;
}
