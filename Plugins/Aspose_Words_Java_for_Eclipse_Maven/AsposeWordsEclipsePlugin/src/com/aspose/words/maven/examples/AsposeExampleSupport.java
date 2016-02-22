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

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import org.apache.commons.io.FileUtils;
import org.eclipse.core.resources.IProject;
import org.eclipse.core.resources.IResource;
import org.eclipse.core.runtime.CoreException;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.w3c.dom.NodeList;
import com.aspose.words.maven.utils.AsposeConstants;
import com.aspose.words.maven.utils.AsposeMavenProjectManager;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;

public class AsposeExampleSupport implements IRunnableWithProgress {

	private String selectedProjectPath;
	private String exampleCategory;
	private IProject project;

	final private static String localExampleFolder = "aspose/GitConsRepos/Aspose.Words/Examples";
	final private static String localExampleSourceFolder = "src/main/java/com/aspose/words/examples";
	final private static String localExampleResourceFolder = "src/main/resources/com/aspose/words/examples";

	public AsposeExampleSupport(String selectedProjectPath, String exampleCategory, IProject project) {
		this.selectedProjectPath = selectedProjectPath;
		this.exampleCategory = exampleCategory;
		this.project = project;
	}

	@Override
	public void run(IProgressMonitor monitor) throws InvocationTargetException, InterruptedException {
		monitor.beginTask("Adding example code in project " + project.getName() + "...", IProgressMonitor.UNKNOWN);
		createExample();
		monitor.done();
	}

	public void createExample() {
		String srcExamplePath = System.getProperty("user.home") + File.separator + localExampleFolder + File.separator
				+ localExampleSourceFolder;
		String srcExampleResourcePath = System.getProperty("user.home") + File.separator + localExampleFolder
				+ File.separator + localExampleResourceFolder;

		String destProjectExamplePath = selectedProjectPath + File.separator + localExampleSourceFolder;
		String destProjectExampleResourcePath = selectedProjectPath + File.separator + localExampleResourceFolder;

		File srcExampleCategoryPath = new File(srcExamplePath + File.separator + exampleCategory);
		File destExampleCategoryPath = new File(destProjectExamplePath + File.separator + exampleCategory);

		Path srcUtil = new File(srcExamplePath + File.separator + "Utils.java").toPath();
		Path destUtil = new File(destProjectExamplePath + File.separator + "Utils.java").toPath();

		File srcExampleResourceCategoryPath = new File(srcExampleResourcePath + File.separator + exampleCategory);
		File destExampleResourceCategoryPath = new File(
				destProjectExampleResourcePath + File.separator + exampleCategory);

		String repositoryPOM_XML = System.getProperty("user.home") + File.separator + localExampleFolder
				+ File.separator + AsposeConstants.MAVEN_POM_XML;

		try {
			FileUtils.copyDirectory(srcExampleCategoryPath, destExampleCategoryPath);
			Files.copy(srcUtil, destUtil, StandardCopyOption.REPLACE_EXISTING);
			FileUtils.copyDirectory(srcExampleResourceCategoryPath, destExampleResourceCategoryPath);

			NodeList examplesNoneAsposeDependencies = AsposeMavenProjectManager.getInstance()
					.getDependenciesFromPOM(repositoryPOM_XML, AsposeConstants.ASPOSE_GROUP_ID);
			AsposeMavenProjectManager.getInstance().addMavenDependenciesInProject(examplesNoneAsposeDependencies);

			NodeList examplesNoneAsposeRepositories = AsposeMavenProjectManager.getInstance()
					.getRepositoriesFromPOM(repositoryPOM_XML, AsposeConstants.ASPOSE_MAVEN_REPOSITORY);
			AsposeMavenProjectManager.getInstance().addMavenRepositoriesInProject(examplesNoneAsposeRepositories);

			project.refreshLocal(IResource.DEPTH_INFINITE, null);

		} catch (IOException | CoreException e) {
			e.printStackTrace();
		}

	}
}