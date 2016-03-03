/*
 * The MIT License (MIT)
 *
 * Copyright (c) 1998-2016 Aspose Pty Ltd.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
package com.aspose.words.maven.examples;

import java.lang.reflect.InvocationTargetException;

import org.eclipse.core.resources.IProject;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.viewers.IStructuredSelection;
import org.eclipse.jface.wizard.Wizard;
import org.eclipse.ui.INewWizard;
import org.eclipse.ui.IWorkbench;

public class AsposeExampleWizard extends Wizard implements INewWizard {

	private AsposeExampleWizardPage wizardPage;

	public AsposeExampleWizard() {
		setWindowTitle("Aspose.Words Code Example");
	}

	@Override
	public void init(IWorkbench workbench, IStructuredSelection selection) {

	}

	@Override
	public void addPages() {
		super.addPages();
		wizardPage = new AsposeExampleWizardPage();
		addPage(wizardPage);
	}

	@Override
	public boolean performFinish() {

		String selectedProjectPath = wizardPage.getSelectedProjectPath();
		String exampleCategory = wizardPage.getSelectedExampleCategory();
		IProject project = wizardPage.getIProject();

		AsposeExampleSupport asposeExampleSupport = new AsposeExampleSupport(selectedProjectPath, exampleCategory,
				project);
		try {
			new ProgressMonitorDialog(this.getShell()).run(true, false, asposeExampleSupport);
		} catch (InvocationTargetException | InterruptedException e) {
			e.printStackTrace();
		}

		return true;
	}

}
