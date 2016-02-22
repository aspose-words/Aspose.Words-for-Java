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
package com.aspose.words.maven.examples;

import java.io.File;
import java.lang.reflect.InvocationTargetException;
import java.net.URI;
import java.util.LinkedList;
import java.util.Queue;
import org.eclipse.core.resources.IProject;
import org.eclipse.core.resources.ResourcesPlugin;
import org.eclipse.core.runtime.CoreException;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.fieldassist.ControlDecoration;
import org.eclipse.jface.fieldassist.FieldDecoration;
import org.eclipse.jface.fieldassist.FieldDecorationRegistry;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.eclipse.jface.wizard.WizardPage;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.widgets.Tree;
import org.eclipse.swt.widgets.TreeItem;
import org.eclipse.swt.widgets.Label;
import org.eclipse.wb.swt.SWTResourceManager;
import com.aspose.words.Activator;
import com.aspose.words.maven.utils.AsposeConstants;
import com.aspose.words.maven.utils.AsposeJavaAPI;
import com.aspose.words.maven.utils.AsposeMavenProjectManager;
import com.aspose.words.maven.utils.AsposeWordsJavaAPI;
import com.aspose.words.maven.utils.FormatExamples;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.ModifyListener;
import org.eclipse.swt.events.ModifyEvent;

public class AsposeExampleWizardPage extends WizardPage {

	private Combo cbProject;
	private Combo cbVersion;
	private Tree examplesTree;

	private ControlDecoration cbProjectDecoration;
	private ControlDecoration cbVersionDecoration;
	private ControlDecoration examplesTreeDecoration;

	/**
	 * Create the wizard.
	 */
	public AsposeExampleWizardPage() {
		super("wizardPage");
		setTitle("New Aspose.Words Code Example");
		setDescription("Aspose.Words Java API - Code Examples");
	}

	public IProject getIProject() {
		return (IProject) cbProject.getData(cbProject.getText());
	}

	public String getSelectedProjectPath() {
		String projectPath = null;
		try {
			IProject project = (IProject) cbProject.getData(cbProject.getText());
			projectPath = project.getDescription().getLocationURI().getPath();

		} catch (CoreException e) {
			e.printStackTrace();
		}
		return projectPath;
	}

	public String getSelectedProjectName() {
		return cbProject.getText();
	}

	public String getSelectedExampleCategory() {
		TreeItem selectedTreeItem = examplesTree.getSelection()[0];
		String selectedCategory = selectedTreeItem.getText().trim().replace(' ', '_').toLowerCase();

		while (selectedTreeItem.getParentItem().getParentItem() != null) {
			selectedTreeItem = selectedTreeItem.getParentItem();
			selectedCategory = selectedTreeItem.getText().trim().replace(' ', '_').toLowerCase() + "/"
					+ selectedCategory;
		}
		return selectedCategory;
	}

	private void downloadExamplesRepo() {
		// download code examples with status progress
		try {
			new ProgressMonitorDialog(this.getShell()).run(true, false, new IRunnableWithProgress() {
				public void run(IProgressMonitor monitor) throws InvocationTargetException, InterruptedException {
					monitor.beginTask("Downloading latest example categories...", IProgressMonitor.UNKNOWN);

					AsposeMavenProjectManager.initialize(null);
					AsposeJavaAPI component = AsposeWordsJavaAPI.initialize(AsposeMavenProjectManager.getInstance());
					component.checkAndUpdateRepo();

					monitor.done();
				}
			});
		} catch (InvocationTargetException | InterruptedException e) {
			e.printStackTrace();
		}
	}

	private void initControls() {

		downloadExamplesRepo();

		IProject[] projects = ResourcesPlugin.getWorkspace().getRoot().getProjects();
		URI path = null;
		try {
			for (IProject project : projects) {
				path = project.getDescription().getLocationURI();
				if (path != null) {
					cbProject.add(project.getDescription().getName());
					cbProject.setData(project.getDescription().getName(), project);
				}
			}
			cbProjectDecoration.hide();
			setPageComplete(true);
			if (cbProject.getItemCount() == 0) {
				cbProject.add(AsposeConstants.API_PROJECT_NOT_FOUND);
				cbProjectDecoration.show();
				setPageComplete(false);
			}

			cbProject.select(0);
		} catch (CoreException e) {
			e.printStackTrace();
		}
	}

	private void initDecorators() {
		FieldDecoration fieldDecoration = FieldDecorationRegistry.getDefault()
				.getFieldDecoration(FieldDecorationRegistry.DEC_ERROR);

		cbProjectDecoration = new ControlDecoration(cbProject, SWT.TOP | SWT.RIGHT);
		cbProjectDecoration.setImage(fieldDecoration.getImage());
		cbProjectDecoration.setDescriptionText("Please first create a Maven project");
		cbProjectDecoration.hide();

		cbVersionDecoration = new ControlDecoration(cbVersion, SWT.TOP | SWT.RIGHT);
		cbVersionDecoration.setImage(fieldDecoration.getImage());
		cbVersionDecoration.setDescriptionText(
				"Please first add maven dependency of " + AsposeConstants.API_NAME + " for java API");
		cbVersionDecoration.hide();

		examplesTreeDecoration = new ControlDecoration(examplesTree, SWT.TOP | SWT.RIGHT);
		examplesTreeDecoration.setImage(fieldDecoration.getImage());
		examplesTreeDecoration.setDescriptionText("Please select one example category");
		examplesTreeDecoration.hide();
	}

	private void onProjectModify() {
		try {
			cbVersion.removeAll();
			IProject selectedProject = (IProject) cbProject.getData(cbProject.getText());
			if (selectedProject != null) {
				String versionNo = AsposeMavenProjectManager.getInstance().getDependencyVersionFromPOM(
						selectedProject.getDescription().getLocationURI(), AsposeConstants.API_MAVEN_DEPENDENCY);
				cbVersionDecoration.hide();
				setPageComplete(true);
				if (versionNo == null) {
					cbVersionDecoration.show();
					setPageComplete(false);
				}
				if (versionNo == null) {
					versionNo = AsposeConstants.API_DEPENDENCY_NOT_FOUND;
				}
				cbVersion.add(versionNo);
				cbVersion.select(0);
			}
		} catch (CoreException e) {
			e.printStackTrace();
		}

	}

	private void onVersionModify() {
		try {
			examplesTree.removeAll();
			if (!cbVersion.getText().equals(AsposeConstants.API_DEPENDENCY_NOT_FOUND)) {
				IProject selectedProject = (IProject) cbProject.getData(cbProject.getText());
				if (selectedProject != null) {
					AsposeMavenProjectManager.initialize(new File(selectedProject.getDescription().getLocationURI()));
					AsposeJavaAPI component = AsposeWordsJavaAPI.initialize(AsposeMavenProjectManager.getInstance());
					populateExamplesTree(component);
					examplesTreeDecoration.show();
					setPageComplete(false);
				}
			}
		} catch (CoreException e) {
			e.printStackTrace();
		}
	}

	private void onTreeSelection() {
		examplesTreeDecoration.show();
		setPageComplete(false);
		TreeItem[] selectedItems = examplesTree.getSelection();
		if (selectedItems[0].getParentItem() != null && selectedItems[0].getItemCount() > 0) {
			examplesTreeDecoration.hide();
			setPageComplete(true);
		}
	}

	/**
	 *
	 * @param asposeComponent
	 * @param top
	 * @param panel
	 */
	public void populateExamplesTree(AsposeJavaAPI asposeComponent) {
		String examplesFullPath = asposeComponent.getLocalRepositoryPath() + File.separator
				+ AsposeConstants.GITHUB_EXAMPLES_SOURCE_LOCATION;
		File directory = new File(examplesFullPath);
		examplesTree.removeAll();
		Queue<Object[]> queue = new LinkedList<>();
		queue.add(new Object[] { null, directory });
		TreeItem top = new TreeItem(examplesTree, 0);
		top.setText("Aspose.Words");

		while (!queue.isEmpty()) {
			Object[] _entry = queue.remove();
			File childFile = ((File) _entry[1]);
			TreeItem parentItem = (TreeItem) _entry[0];
			if (childFile.isDirectory()) {
				if (parentItem != null) {
					TreeItem child = new TreeItem(parentItem, SWT.NONE);
					child.setText(FormatExamples.formatTitle(childFile.getName()));
					parentItem = child;
				} else {
					parentItem = top;
				}
				for (File f : childFile.listFiles()) {
					queue.add(new Object[] { parentItem, f });
				}
			} else if (childFile.isFile()) {
				TreeItem child = new TreeItem(parentItem, SWT.NONE);
				child.setText(FormatExamples.formatTitle(childFile.getName()));
			}
		}
	}

	/**
	 * Create contents of the wizard.
	 * 
	 * @param parent
	 */
	public void createControl(Composite parent) {
		Composite container = new Composite(parent, SWT.NULL);

		setControl(container);

		cbProject = new Combo(container, SWT.READ_ONLY);
		cbProject.addModifyListener(new ModifyListener() {
			public void modifyText(ModifyEvent e) {
				onProjectModify();
			}
		});
		cbProject.setBounds(181, 101, 366, 23);

		examplesTree = new Tree(container, SWT.BORDER);
		examplesTree.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				onTreeSelection();
			}
		});
		examplesTree.setBounds(27, 170, 520, 140);

		Label label = new Label(container, SWT.NONE);
		label.setImage(SWTResourceManager.getImage(Activator.getResourceFilePath("long_banner.png")));
		label.setBounds(10, 0, 564, 95);

		Label lblProject = new Label(container, SWT.NONE);
		lblProject.setBounds(134, 104, 40, 15);
		lblProject.setText("Project:");

		Label lblAsposewordsForJava = new Label(container, SWT.NONE);
		lblAsposewordsForJava.setBounds(5, 131, 170, 15);
		lblAsposewordsForJava.setText("Aspose.Words for Java (version):");

		cbVersion = new Combo(container, SWT.READ_ONLY);
		cbVersion.addModifyListener(new ModifyListener() {
			public void modifyText(ModifyEvent e) {
				onVersionModify();
			}
		});
		cbVersion.setBounds(181, 128, 366, 23);

		initDecorators();
		initControls();
	}
}
