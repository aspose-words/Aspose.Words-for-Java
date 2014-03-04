/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
import com.aspose.words.*;

import javax.swing.*;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
* Shows an About form for the DocumentExplorer application.
*/
public class AboutForm extends JDialog
{
	private JPanel contentPane;
	private JButton buttonOK;
	private JTextPane textPane1;
	private JTextPane textPane2;
	private JLabel asposeIcon;

	public AboutForm()
	{
		setResizable(false);
		textPane1.setText("Use DocumentExplorer to:\n" +
		                  "-   Learn from source code how to use Aspose.Words in your project.\n" +
		                  "-   Visually explore document elements in the Aspose.Words Object Model.\n" +
		                  "using Aspose.Words.\n" +
                          "-   Quickly convert between DOC, DOCX, ODF, EPUB, PDF, RTF, SWF, WordML, HTML, " +
                          "MHTML and plain text formats.\n");
		textPane2.setFont(new Font("Verdana", Font.BOLD, 18));
		asposeIcon.setIcon(Utils.createImageIcon("images/aspose.gif"));

		setContentPane(contentPane);
		setModal(true);
		getRootPane().setDefaultButton(buttonOK);

		buttonOK.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				onOK();
			}
		});
	}

	private void onOK()
	{
		dispose();
	}

}