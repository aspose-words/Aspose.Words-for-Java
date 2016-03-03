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
package com.aspose.words.maven;

import java.lang.reflect.InvocationTargetException;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.viewers.IStructuredSelection;
import org.eclipse.jface.wizard.Wizard;
import org.eclipse.ui.INewWizard;
import org.eclipse.ui.IWorkbench;

public class AsposeMavenProjectWizard extends Wizard implements INewWizard {

	private AsposeMavenProjectWizardPage wizardPage;

	public AsposeMavenProjectWizard() {
		setWindowTitle("Aspose.Words Maven Project");
	}

	@Override
	public void init(IWorkbench workbench, IStructuredSelection selection) {

	}

	@Override
	public void addPages() {
		super.addPages();
		wizardPage = new AsposeMavenProjectWizardPage();
		addPage(wizardPage);
	}

	@Override
	public boolean performFinish() {
		AsposeMavenProjectSupport asposeMavenProjectSupport = new AsposeMavenProjectSupport(wizardPage.getProjectName(),
				wizardPage.getLocationURI(), wizardPage.getPackageName(), wizardPage.isDownloadExamplesChecked(),
				wizardPage.getVersion(), wizardPage.getGroupId());
		try {
			new ProgressMonitorDialog(this.getShell()).run(true, false, asposeMavenProjectSupport);
		} catch (InvocationTargetException | InterruptedException e) {
			e.printStackTrace();
		}
		return true;
	}

}
