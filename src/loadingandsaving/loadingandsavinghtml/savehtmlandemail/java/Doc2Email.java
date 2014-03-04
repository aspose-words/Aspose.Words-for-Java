//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package loadingandsaving.loadingandsavinghtml.savehtmlandemail.java;

import com.aspose.words.*;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import com.sun.mail.smtp.SMTPTransport;

import javax.swing.*;
import javax.swing.text.JTextComponent;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.text.MessageFormat;
import java.util.Date;
import java.util.Properties;

public class Doc2Email extends JFrame {
    private JPanel contentPane;
    private JButton buttonOpen;
    private JButton buttonSend;
    private JTextPane textSmtp;
    private JTextPane textEmail;
    private JTextPane textRecipient;
    private JTextPane textSubject;
    private JCheckBox secureConnectionCheckBox;
    private JPasswordField passwordField;
    private JFormattedTextField formattedPortField;
    private JLabel labelSentMessage;

    private String mDocumentPath = Doc2Email.class.getProtectionDomain().getCodeSource().getLocation().getPath();
    private String mDocumentName;
    private Document mDocument;

    private String APPLICATION_TITLE = "Doc2Email";

    // Open File filters
    static final OpenFileFilter OPEN_FILE_FILTER_ALL_SUPPORTED_FORMATS = new OpenFileFilter(
            new String[] {".doc",".dot",".docx;",".dotx;",".docm",".dotm",".xml",".wml",".rtf",".odt",".ott",".htm",".html",".xhtml",".mht",".mhtm",".mhtml"}, "All Supported Formats (*.doc;*.dot;*.docx;*.dotx;*.docm;*.dotm;*.xml;*.wml;*.rtf;*.odt;*.ott;*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml)");

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

    public Doc2Email() {
        // Setup the main window
        setContentPane(contentPane);
        getRootPane().setDefaultButton(buttonOpen);

        pack();
        setTitle("Doc2Email");
        setLocationRelativeTo(null);
        disableTextFields();
        setVisible(true);

        buttonOpen.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onOpen();
            }
        });

        buttonSend.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onSend();
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
     * Called when user presses the cross on the window or the escape key.
     */
    private void closeWindow()
    {
        System.exit(0);
    }

    /**
     * Called when the user presses the "Open" button. Displays a dialog which allows the user to open a document
     * using Aspose.Words.
     */
    private void onOpen() {

        try
        {
            // Display dialog and find the desired path.
            String fileName = openDocument();

            labelSentMessage.setVisible(false);
            contentPane.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

            if (!"".equals(fileName)) {
                // Load the document.
                mDocument = new Document(fileName);
                setTitle(APPLICATION_TITLE + " - " + mDocumentName);
                enableTextFields();
            }

        }

        catch (Exception e)
        {
            JOptionPane.showMessageDialog(this,
                    MessageFormat.format("An error occured while opening the document: {0}", e.getMessage()),
                    APPLICATION_TITLE, JOptionPane.ERROR_MESSAGE);
        }

        finally
        {
            contentPane.setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }

    }

    /**
     * Displays a file chooser dialog.
     */
    private String openDocument()
    {
        // Create the file chooser and set the appropriate filters.
        JFileChooser mOpenDialog = new JFileChooser();
        mOpenDialog.setAcceptAllFileFilterUsed(false);
        mOpenDialog.setFileFilter(OPEN_FILE_FILTER_DOC_FORMAT);
        mOpenDialog.setFileFilter(OPEN_FILE_FILTER_DOCX_FORMAT);
        mOpenDialog.setFileFilter(OPEN_FILE_FILTER_XML_FORMAT);
        mOpenDialog.setFileFilter(OPEN_FILE_FILTER_RTF_FORMAT);
        mOpenDialog.setFileFilter(OPEN_FILE_FILTER_ODT_FORMAT);
        mOpenDialog.setFileFilter(OPEN_FILE_FILTER_HTML_FORMAT);
        mOpenDialog.setFileFilter(OPEN_FILE_FILTER_ALL_SUPPORTED_FORMATS); // This is last so it will appear by default.
        mOpenDialog.setMultiSelectionEnabled(false);
        mOpenDialog.setFileSelectionMode(JFileChooser.FILES_ONLY);
        mOpenDialog.setDialogTitle("Open Document");

        mOpenDialog.setCurrentDirectory(new File(mDocumentPath));

        if (mOpenDialog.showOpenDialog(this) == JFileChooser.APPROVE_OPTION)
        {
            File file = mOpenDialog.getSelectedFile();
            String fileName = file.getAbsolutePath();
            if (file.exists())
            {
                mDocumentPath = file.getParent();
                mDocumentName = file.getName();
                return fileName;
            }
            else
            {
                JOptionPane.showMessageDialog(this,
                        MessageFormat.format("File \"{0}\" doesn't exist.", fileName),
                        APPLICATION_TITLE, JOptionPane.ERROR_MESSAGE);
                return "";
            }
        }
        else
        {
            return "";
        }
    }

    /**
     * Called when the user presses the "Send" button. Verifies that all fields are filled out and passes this information to the
     * send method to dispatch the mail.
     */
    private void onSend() {

        if(textSmtp.getText().length() == 0 || textEmail.getText().length() == 0 || new String(passwordField.getPassword()).length() == 0
                || textRecipient.getText().length() == 0 || textSubject.getText().length() == 0 || formattedPortField.getText().trim().length() == 0)
        {
            JOptionPane.showMessageDialog(this,
                    "All fields are required. Please enter the appropriate information.",
                    APPLICATION_TITLE, JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        // Set these as disabled during sending.
        disableTextFields();
        labelSentMessage.setVisible(false);
        contentPane.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        buttonSend.setEnabled(false);

        // This is required to get the panel to refreshing the above changes before sending the message.
        this.update(getGraphics());

        try
        {
            send(textSmtp.getText(), textEmail.getText(), new String(passwordField.getPassword()), textRecipient.getText(), textSubject.getText(), Integer.parseInt(formattedPortField.getText().trim()), secureConnectionCheckBox.isSelected());
            labelSentMessage.setVisible(true);
        }

        catch(Exception e)
        {
            JOptionPane.showMessageDialog(this,
                    MessageFormat.format("An error occured while sending: {0}.", e.getMessage()),
                    APPLICATION_TITLE, JOptionPane.ERROR_MESSAGE);
        }

        finally
        {
            buttonSend.setEnabled(true);
            enableTextFields();
            contentPane.setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    /**
     * Disables the text fields used to enter in mailing information.
     */
    private void disableTextFields()
    {
        disableTextField(textSmtp);
        disableTextField(textEmail);
        disableTextField(textRecipient);
        disableTextField(textSubject);
        disableTextField(passwordField);
        disableTextField(formattedPortField);
        secureConnectionCheckBox.setEnabled(false);
        buttonSend.setEnabled(false);
    }

    /**
     * Disables a text field.
     */
    private void disableTextField(JTextComponent text)
    {
        text.setBorder(BorderFactory.createEtchedBorder());
        text.setEditable(false);
        text.setBackground(UIManager.getColor("TextField.disabled"));
    }

    /**
     * Enables the text fields used to enter in mailing information for use.
     */
    private void enableTextFields()
    {
        enableTextField(textSmtp);
        enableTextField(textEmail);
        enableTextField(textRecipient);
        enableTextField(textSubject);
        enableTextField(passwordField);
        enableTextField(formattedPortField);
        secureConnectionCheckBox.setEnabled(true);
        buttonSend.setEnabled(true);
    }

    /**
     * Enables a specific text field.
     */
    private void enableTextField(JTextComponent text)
    {
        text.setEditable(true);
        text.setBackground(UIManager.getColor("TextField.background"));
    }

    /**
     * Convert document to HTML mail message and send it to recipient
     *
     * @param smtp Smtp server.
     * @param emailFrom Sender's e-mail.
     * @param password Sender password.
     * @param emailTo Recipient e-mail.
     * @param subject E-mail subject.
     * @param port The port to send on.
     * @param secureConnection Sets if authentication is used.
     */
    private void send(String smtp, String emailFrom, String password, String emailTo, String subject, int port, boolean secureConnection) throws Exception
    {
        // Create a temporary directory where any images of the exported document are stored.
        File tempDir = new File(System.getProperty("java.io.tmpdir") + "AsposeMail\\");

        if (!tempDir.exists())
            tempDir.mkdir();

        // Save the document in HTML format.
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        // Save the images in the temporary folder.
        saveOptions.setImagesFolder(tempDir.getAbsolutePath());
        // We want the images in the HTML to be referenced in the e-mail as attachments so add the cid prefix to the image file name.
        // This replaces what would be the path to the image with the "cid" prefix.
        saveOptions.setImagesFolderAlias("cid:");
        // Headers and footers normally don't export well to HTML so disable them.
        saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.NONE);
        mDocument.save(output, saveOptions);
        // Get the HTML as a string.
        String htmlMessage = output.toString("UTF-8");

        // Setup the mail client.
        Properties props = System.getProperties();
        props.put("mail.smtps.host",smtp);
        props.put("mail.smtps.auth", secureConnection ? "true" : "false");
        props.put("mail.smtp.port", port);
        Session session = Session.getInstance(props, null);
        Message message = new MimeMessage(session);
        message.setFrom(new InternetAddress(emailFrom));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(emailTo, false));
        message.setSubject(subject);

        // The body is made up of the HTML content along with the images as embedded attachments.
        Multipart mp = new MimeMultipart();
        MimeBodyPart htmlPart = new MimeBodyPart();
        htmlPart.setContent(htmlMessage, "text/html");

        mp.addBodyPart(htmlPart);

        for(File file : tempDir.listFiles())
        {
            MimeBodyPart imagePart = new MimeBodyPart();
            DataSource fds = new FileDataSource(file.getAbsolutePath());
            imagePart.setDataHandler(new DataHandler(fds));
            // Content-ID should match the ID after the cid prefix. In this case this will be the filename with a slash
            // e.g /Aspose.Words.57647.png
            imagePart.setHeader("Content-ID", MessageFormat.format("</{0}>", file.getName()));
            mp.addBodyPart(imagePart);
        }

        // Attach the content to the message.
        message.setContent(mp);

        message.setSentDate(new Date());
        SMTPTransport transport = (SMTPTransport)session.getTransport("smtps");
        transport.connect(smtp, emailFrom, password);
        transport.sendMessage(message, message.getAllRecipients());
        transport.close();

        // Clean up the temporary files and remove the directory.
        for(File file : tempDir.listFiles())
            file.delete();

        tempDir.delete();
    }

     /**
     * Called to create a UI component manually.
     */
    private void createUIComponents() {
        try{
            // We want our own mask formatter to match any number of numeric characters up to a length of five digits long.
            VariableLengthMaskFormatter formatter = new VariableLengthMaskFormatter("#####");
            formattedPortField = new JFormattedTextField(formatter);
        }

        catch(Exception e)
        {
        }
    }
}