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

import java.net.URI;
import java.nio.file.Files;
import org.eclipse.core.resources.ICommand;
import org.eclipse.core.resources.IContainer;
import org.eclipse.core.resources.IFolder;
import org.eclipse.core.resources.IProject;
import org.eclipse.core.resources.IProjectDescription;
import org.eclipse.core.resources.IResource;
import org.eclipse.core.resources.ResourcesPlugin;
import org.eclipse.core.runtime.CoreException;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.jface.operation.IRunnableWithProgress;
import com.aspose.words.Activator;
import com.aspose.words.maven.utils.AsposeJavaAPI;
import com.aspose.words.maven.utils.AsposeMavenProjectManager;
import com.aspose.words.maven.utils.AsposeWordsJavaAPI;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;

public class AsposeMavenProjectSupport implements IRunnableWithProgress {

	private String projectName;
	private URI location;
	private String packageName;
	private boolean downloadExamples;
	private String version;
	private String groupId;

	public AsposeMavenProjectSupport(String projectName, URI location, String packageName, boolean downloadExamples,
			String version, String groupId) {
		this.projectName = projectName;
		this.location = location;
		this.packageName = packageName;
		this.downloadExamples = downloadExamples;
		this.version = version;
		this.groupId = groupId;
	}

	@Override
	public void run(IProgressMonitor monitor) throws InvocationTargetException, InterruptedException {
		monitor.beginTask("Processing...", IProgressMonitor.UNKNOWN);
		createProject(monitor);
		monitor.done();
	}

	private IProject createProject(IProgressMonitor monitor) {
		IProject project = createBaseProject(projectName, location);
		try {
			monitor.setTaskName("Creating project...");
			String[] paths = { "src/main/java/" + packageName.replace(".", "/"), "src/test/java", "src/main/resources" };
			addToProjectStructure(project, paths);

			Files.copy(new File(Activator.getResourceFilePath("pom-xml-template.txt")).toPath(),
					new File(location.getPath() + "/pom.xml").toPath());
			Files.copy(new File(Activator.getResourceFilePath("classpath-template.txt")).toPath(),
					new File(location.getPath() + "/.classpath").toPath());

			Files.createDirectories(new File(location.getPath() + "/.settings").toPath());
			Files.copy(new File(Activator.getResourceFilePath("org-eclipse-jdt-core.txt")).toPath(),
					new File(location.getPath() + "/.settings/org.eclipse.jdt.core.prefs").toPath());

			AsposeMavenProjectManager asposeMavenProjectManager = AsposeMavenProjectManager
					.initialize(new File(location));
			asposeMavenProjectManager.configureProjectMavenPOM(groupId, projectName, version);
			project.refreshLocal(IResource.DEPTH_INFINITE, null);

			if (downloadExamples) {
				monitor.setTaskName("Downloading code examples...");
				AsposeMavenProjectManager.initialize(new File(location));
				AsposeJavaAPI component = AsposeWordsJavaAPI.initialize(AsposeMavenProjectManager.getInstance());
				component.checkAndUpdateRepo();
			}

		} catch (CoreException e) {
			e.printStackTrace();
			project = null;
		} catch (IOException e) {
			e.printStackTrace();
		}

		return project;
	}

	private IProject createBaseProject(String projectName, URI location) {
		IProject newProject = ResourcesPlugin.getWorkspace().getRoot().getProject(projectName);
		String natures[] = { "org.eclipse.jdt.core.javanature", "org.eclipse.m2e.core.maven2Nature" };

		if (!newProject.exists()) {
			URI projectLocation = location;
			IProjectDescription desc = newProject.getWorkspace().newProjectDescription(newProject.getName());

			ICommand commandJavaBuilder = (ICommand) desc.newCommand();
			ICommand commandMaven2Builder = (ICommand) desc.newCommand();
			commandJavaBuilder.setBuilderName("org.eclipse.jdt.core.javabuilder");
			commandMaven2Builder.setBuilderName("org.eclipse.m2e.core.maven2Builder");
			ICommand buildspecs[] = { commandJavaBuilder, commandMaven2Builder };

			desc.setBuildSpec(buildspecs);
			desc.setNatureIds(natures);

			if (location != null && ResourcesPlugin.getWorkspace().getRoot().getLocationURI().equals(location)) {
				projectLocation = null;
			}

			desc.setLocationURI(projectLocation);
			try {
				newProject.create(desc, null);
				if (!newProject.isOpen()) {
					newProject.open(null);
				}
			} catch (CoreException e) {
				e.printStackTrace();
			}
		}

		return newProject;
	}

	private void createFolder(IFolder folder) throws CoreException {
		IContainer parent = folder.getParent();
		if (parent instanceof IFolder) {
			createFolder((IFolder) parent);
		}
		if (!folder.exists()) {
			folder.create(false, true, null);
		}
	}

	/**
	 * Create a folder structure with a parent root, overlay, and a few child
	 * folders.
	 *
	 * @param newProject
	 * @param paths
	 * @throws CoreException
	 */
	private void addToProjectStructure(IProject newProject, String[] paths) throws CoreException {
		for (String path : paths) {
			IFolder etcFolders = newProject.getFolder(path);
			createFolder(etcFolders);
		}
	}

}