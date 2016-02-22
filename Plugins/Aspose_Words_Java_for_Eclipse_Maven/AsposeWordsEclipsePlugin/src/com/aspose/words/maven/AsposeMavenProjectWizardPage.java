/*
 * The MIT License (MIT)
 *
 * Copyright (c) 1998-2015 Aspose Pty Ltd.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
package com.aspose.words.maven;

import java.io.File;
import java.net.URI;
import org.eclipse.core.resources.IProject;
import org.eclipse.core.resources.ResourcesPlugin;
import org.eclipse.core.runtime.CoreException;
import org.eclipse.jface.wizard.WizardPage;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.DirectoryDialog;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;
import org.eclipse.wb.swt.SWTResourceManager;
import com.aspose.words.Activator;
import com.aspose.words.maven.utils.MavenSettings;
import org.eclipse.swt.events.ModifyListener;
import org.eclipse.swt.events.ModifyEvent;
import org.eclipse.jface.fieldassist.ControlDecoration;
import org.eclipse.jface.fieldassist.FieldDecoration;
import org.eclipse.jface.fieldassist.FieldDecorationRegistry;

public class AsposeMavenProjectWizardPage extends WizardPage {

	public static final String PROP_PROJECT_NAME = "projectName";
	public static final String PROP_GROUP_ID = "groupId";

	private Text txtProjectLocation;
	private Text txtProjectName;
	private Text txtProjectFolder;
	private Text txtArtifactId;
	private Text txtGroupId;
	private Text txtVersion;
	private Text txtPackage;
	private Button chkDownloadExamples;

	private ControlDecoration txtProjectNameDecoration;
	private ControlDecoration txtProjectLocationDecoration;
	private ControlDecoration txtGroupIdDecoration;
	private ControlDecoration txtVersionDecoration;
	private ControlDecoration txtProjectFolderDecoration;
	private ControlDecoration txtPackageDecoration;

	/**
	 * Create the wizard.
	 */
	public AsposeMavenProjectWizardPage() {
		super("wizardPage");
		setTitle("New Project");
		setDescription("Name and Location");
	}

	private String getDefaultProjectName() {
		String defaultName = "asposemavenproject";
		IProject[] projects = ResourcesPlugin.getWorkspace().getRoot().getProjects();
		try {
			for (int i = 1; i < 100; i++) {

				boolean match = false;
				for (IProject project : projects) {
					if (project.getDescription().getName().equals(defaultName + i)) {
						match = true;
						break;
					}

				}
				if (!match) {
					defaultName = defaultName + i;
					break;
				}
			}
		} catch (CoreException e) {
			e.printStackTrace();
		}
		return defaultName;
	}

	private void initControls() {
		txtGroupId.setText(MavenSettings.getDefault().getLastArchetypeGroupId());
		txtVersion.setText(MavenSettings.getDefault().getLastArchetypeVersion());
		txtProjectName.setText(getDefaultProjectName());
		txtProjectName.setSelection(txtProjectName.getCharCount());
		txtProjectLocation.setText(ResourcesPlugin.getWorkspace().getRoot().getLocation().toOSString());
		txtProjectFolder.setText(txtProjectLocation.getText() + File.separator + txtProjectName.getText());
		txtArtifactId.setText(txtProjectName.getText());
		txtPackage.setText(txtGroupId.getText() + "." + txtProjectName.getText());
	}

	private void initDecorators() {
		FieldDecoration fieldDecoration = FieldDecorationRegistry.getDefault()
				.getFieldDecoration(FieldDecorationRegistry.DEC_ERROR);

		txtProjectNameDecoration = new ControlDecoration(txtProjectName, SWT.TOP | SWT.RIGHT);
		txtProjectNameDecoration.setImage(fieldDecoration.getImage());
		txtProjectNameDecoration.hide();

		txtProjectLocationDecoration = new ControlDecoration(txtProjectLocation, SWT.TOP | SWT.RIGHT);
		txtProjectLocationDecoration.setImage(fieldDecoration.getImage());
		txtProjectLocationDecoration.hide();

		txtGroupIdDecoration = new ControlDecoration(txtGroupId, SWT.TOP | SWT.RIGHT);
		txtGroupIdDecoration.setImage(fieldDecoration.getImage());
		txtGroupIdDecoration.hide();

		txtVersionDecoration = new ControlDecoration(txtVersion, SWT.TOP | SWT.RIGHT);
		txtVersionDecoration.setImage(fieldDecoration.getImage());
		txtVersionDecoration.hide();

		txtProjectFolderDecoration = new ControlDecoration(txtProjectFolder, SWT.TOP | SWT.RIGHT);
		txtProjectFolderDecoration.setImage(fieldDecoration.getImage());
		txtProjectFolderDecoration.hide();

		txtPackageDecoration = new ControlDecoration(txtPackage, SWT.TOP | SWT.RIGHT);
		txtPackageDecoration.setImage(fieldDecoration.getImage());
		txtPackageDecoration.hide();
	}

	private void onProjectNameChange() {
		txtProjectNameDecoration.hide();
		setPageComplete(true);
		if (txtProjectName.getText().trim().length() == 0) {
			txtProjectNameDecoration.setDescriptionText("Project Name is not a valid folder name");
			txtProjectNameDecoration.show();
			setPageComplete(false);
		}
		txtProjectFolder.setText(txtProjectLocation.getText() + File.separator + txtProjectName.getText());
		txtArtifactId.setText(txtProjectName.getText());
		txtPackage.setText("com.mycompany." + txtProjectName.getText());
	}

	private void onProjectLocationChange() {
		txtProjectLocationDecoration.hide();
		setPageComplete(true);
		if (!new File(txtProjectLocation.getText().trim()).isDirectory()) {
			txtProjectLocationDecoration.setDescriptionText("Project Folder is not a valid path");
			txtProjectLocationDecoration.show();
			setPageComplete(false);
		}
		txtProjectFolder.setText(txtProjectLocation.getText() + File.separator + txtProjectName.getText());
	}

	private void onVersionChange() {
		txtVersionDecoration.hide();
		setPageComplete(true);
		if (txtVersion.getText().trim().length() == 0) {
			txtVersionDecoration.setDescriptionText("Version may not be empty");
			txtVersionDecoration.show();
			setPageComplete(false);
		}
	}

	private void onPackageChange() {
		txtPackageDecoration.hide();
		setPageComplete(true);
		String packageName = txtPackage.getText().trim();
		if (!(packageName.equals("")
				|| packageName.matches("([\\p{L}_$][\\p{L}\\p{N}_$]*\\.)*[\\p{L}_$][\\p{L}\\p{N}_$]*"))) {
			txtPackageDecoration.setDescriptionText("Package may not be empty");
			txtPackageDecoration.show();
			setPageComplete(false);
		}
	}

	private void onGroupIdChange() {
		txtGroupIdDecoration.hide();
		setPageComplete(true);
		if (txtGroupId.getText().trim().length() == 0) {
			txtGroupIdDecoration.setDescriptionText("GroupdId may not be empty");
			txtGroupIdDecoration.show();
			setPageComplete(false);
		}
		txtPackage.setText(txtGroupId.getText() + "." + txtProjectName.getText());
	}

	private void onProjectFolderChange() {
		txtProjectFolderDecoration.hide();
		setPageComplete(true);
		File projLoc = new File(
				(new File(txtProjectLocation.getText()).getAbsoluteFile()).toURI().normalize().getPath());
		File destFolder = new File(
				(new File(txtProjectFolder.getText()).getAbsoluteFile()).toURI().normalize().getPath());

		while (projLoc != null && !projLoc.exists()) {
			projLoc = projLoc.getParentFile();
		}
		if (projLoc == null || !projLoc.canWrite()) {
			txtProjectFolderDecoration.setDescriptionText("Project Folder cannot be created");
			txtProjectFolderDecoration.show();
			setPageComplete(false);
		} else {
			File[] kids = destFolder.listFiles();
			if (destFolder.exists() && kids != null && kids.length > 0) {
				txtProjectFolderDecoration.setDescriptionText("Project Folder already exists and is not empty");
				txtProjectFolderDecoration.show();
				setPageComplete(false);
			}
		}
	}

	public String getProjectName() {
		return txtProjectName.getText();
	}

	public String getPackageName() {
		return txtPackage.getText();
	}

	public URI getLocationURI() {
		return new File(txtProjectFolder.getText()).toURI();
	}

	public String getVersion() {
		return txtVersion.getText();
	}

	public String getGroupId() {
		return txtGroupId.getText();
	}

	public boolean isDownloadExamplesChecked() {
		return chkDownloadExamples.getSelection();
	}

	/**
	 * Create contents of the wizard.
	 * 
	 * @param parent
	 */
	public void createControl(Composite parent) {
		Composite container = new Composite(parent, SWT.NULL);

		setControl(container);

		Label lblNewLabel = new Label(container, SWT.NONE);
		lblNewLabel.setImage(SWTResourceManager.getImage(Activator.getResourceFilePath("long_banner.png")));
		lblNewLabel.setBounds(10, 0, 564, 95);

		Label lblPleaseEnterProject = new Label(container, SWT.NONE);
		lblPleaseEnterProject.setFont(SWTResourceManager.getFont("Segoe UI", 9, SWT.BOLD));
		lblPleaseEnterProject.setBounds(5, 98, 179, 15);
		lblPleaseEnterProject.setText("Please enter project detail:");

		Label lblPleaseEnterMaven = new Label(container, SWT.NONE);
		lblPleaseEnterMaven.setText("Please enter maven artifact detail:");
		lblPleaseEnterMaven.setFont(SWTResourceManager.getFont("Segoe UI", 9, SWT.BOLD));
		lblPleaseEnterMaven.setBounds(5, 194, 213, 15);

		Label lblProjectName = new Label(container, SWT.NONE);
		lblProjectName.setBounds(5, 118, 84, 15);
		lblProjectName.setText("Project Name:");

		Label lblProjectLocation = new Label(container, SWT.NONE);
		lblProjectLocation.setBounds(5, 146, 94, 15);
		lblProjectLocation.setText("Project Location:");

		Label lblProjectFolder = new Label(container, SWT.NONE);
		lblProjectFolder.setText("Project Folder:");
		lblProjectFolder.setBounds(5, 173, 94, 15);

		Label lblArtifactId = new Label(container, SWT.NONE);
		lblArtifactId.setText("Artifact Id:");
		lblArtifactId.setBounds(5, 215, 94, 15);

		Label lblGroupId = new Label(container, SWT.NONE);
		lblGroupId.setText("Group Id:");
		lblGroupId.setBounds(5, 242, 94, 15);

		Label lblVersion = new Label(container, SWT.NONE);
		lblVersion.setText("Version:");
		lblVersion.setBounds(5, 270, 94, 15);

		Label lblPackage = new Label(container, SWT.NONE);
		lblPackage.setText("Package:");
		lblPackage.setBounds(5, 296, 94, 15);

		txtProjectName = new Text(container, SWT.BORDER);
		txtProjectName.addModifyListener(new ModifyListener() {
			public void modifyText(ModifyEvent e) {
				onProjectNameChange();
			}
		});
		txtProjectName.setBounds(118, 116, 370, 21);

		txtProjectLocation = new Text(container, SWT.BORDER);
		txtProjectLocation.addModifyListener(new ModifyListener() {
			public void modifyText(ModifyEvent e) {
				onProjectLocationChange();
			}
		});
		txtProjectLocation.setBounds(118, 143, 370, 21);

		Button btnNewButton = new Button(container, SWT.NONE);
		btnNewButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				DirectoryDialog dialog = new DirectoryDialog(new Shell(), SWT.NULL);
				String path = dialog.open();
				if (path != null) {
					txtProjectLocation.setText(path);
				}
			}
		});
		btnNewButton.setBounds(494, 140, 75, 25);
		btnNewButton.setText("Browse...");

		txtProjectFolder = new Text(container, SWT.BORDER);
		txtProjectFolder.addModifyListener(new ModifyListener() {
			public void modifyText(ModifyEvent e) {
				onProjectFolderChange();
			}
		});
		txtProjectFolder.setEditable(false);
		txtProjectFolder.setBounds(118, 170, 370, 21);

		txtArtifactId = new Text(container, SWT.BORDER);
		txtArtifactId.setEnabled(false);
		txtArtifactId.setEditable(false);
		txtArtifactId.setBounds(118, 212, 370, 21);

		txtGroupId = new Text(container, SWT.BORDER);
		txtGroupId.addModifyListener(new ModifyListener() {
			public void modifyText(ModifyEvent e) {
				onGroupIdChange();
			}
		});
		txtGroupId.setBounds(118, 239, 370, 21);

		txtVersion = new Text(container, SWT.BORDER);
		txtVersion.addModifyListener(new ModifyListener() {
			public void modifyText(ModifyEvent e) {
				onVersionChange();
			}
		});
		txtVersion.setBounds(118, 266, 370, 21);

		txtPackage = new Text(container, SWT.BORDER);
		txtPackage.addModifyListener(new ModifyListener() {
			public void modifyText(ModifyEvent e) {
				onPackageChange();
			}
		});
		txtPackage.setBounds(118, 293, 370, 21);

		chkDownloadExamples = new Button(container, SWT.CHECK);
		chkDownloadExamples.setBounds(5, 325, 492, 16);
		chkDownloadExamples.setText("Also Download Code Examples (for using Aspose.Words for Java)");

		Label lblNewLabel_1 = new Label(container, SWT.NONE);
		lblNewLabel_1.setBounds(494, 296, 55, 15);
		lblNewLabel_1.setText("(Optional)");

		initDecorators();
		initControls();
	}
}
