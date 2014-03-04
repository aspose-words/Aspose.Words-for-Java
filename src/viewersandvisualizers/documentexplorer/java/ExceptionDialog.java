/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
import com.aspose.words.*;

import javax.swing.*;
import java.awt.event.*;

/**
* Provides information to the user if an an unexepected exception occurs.
*/
public class ExceptionDialog extends JDialog
{
	private JPanel contentPane;
	private JButton buttonOK;
	private JButton buttonCancel;
	private JTextPane text1;

	void InitializeComponent()
	{
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

		setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
		addWindowListener(new WindowAdapter()
		{
			public void windowClosing(WindowEvent e)
			{
				onOK();
			}
		});

		contentPane.registerKeyboardAction(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				onOK();
			}
		}, KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), JComponent.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT);
	}

	public ExceptionDialog()
	{
		InitializeComponent();
	}

	private void onOK()
	{
		dispose();
	}

	public ExceptionDialog(Exception ex)
	{
		InitializeComponent();
		this.setTitle(Globals.UNEXPECTED_EXCEPTION_DIALOG_TITLE);

		this.text1.setText("\r\n" + ex.toString() + "\r\n");
		this.text1.setSelectionStart(this.text1.getText().length());
	}

}