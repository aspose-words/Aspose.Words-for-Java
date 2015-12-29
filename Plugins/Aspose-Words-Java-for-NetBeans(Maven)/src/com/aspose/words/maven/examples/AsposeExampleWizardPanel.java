/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.aspose.words.maven.examples;

import com.aspose.words.maven.utils.AsposeConstants;
import com.aspose.words.maven.utils.AsposeMavenProjectManager;
import com.aspose.words.maven.utils.AsposeWordsJavaAPI;
import java.io.File;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import org.netbeans.api.project.Project;
import org.netbeans.spi.project.ui.templates.support.Templates;
import org.openide.WizardDescriptor;
import org.openide.util.Exceptions;
import org.openide.util.HelpCtx;
import org.w3c.dom.NodeList;

/**
 * @author Adeel Ilyas
 */
public class AsposeExampleWizardPanel implements WizardDescriptor.Panel<WizardDescriptor> {

    private AsposeExamplePanel component;
    private static boolean storeSettingsCalled = false;

    /**
     *
     * @return
     */
    @Override
    public AsposeExamplePanel getComponent() {
        if (component == null) {
            component = new AsposeExamplePanel(this);
        }
        return component;
    }

    /**
     *
     * @return
     */
    @Override
    public HelpCtx getHelp() {

        return HelpCtx.DEFAULT_HELP;

    }

    /**
     *
     * @return
     */
    @Override
    public boolean isValid() {
        // Enables Finish / OK /Next button 
        return component.validateDialog();

    }

    private final Set<ChangeListener> listeners = new HashSet<>(1); // or can use ChangeSupport in NB 6.0

    /**
     *
     * @param l
     */
    @Override
    public void addChangeListener(ChangeListener l) {
        synchronized (listeners) {
            listeners.add(l);
        }
    }

    /**
     *
     * @param l
     */
    @Override
    public void removeChangeListener(ChangeListener l) {
        synchronized (listeners) {
            listeners.remove(l);
        }
    }

    /**
     *
     */
    protected final void fireChangeEvent() {

        Iterator<ChangeListener> it;
        synchronized (listeners) {
            it = new HashSet<>(listeners).iterator();
        }
        ChangeEvent ev = new ChangeEvent(this);
        while (it.hasNext()) {
            it.next().stateChanged(ev);
        }
    }

    /**
     *
     * @param wiz
     */
    @Override
    public void readSettings(WizardDescriptor wiz) {
        Project selectedProject = Templates.getProject(wiz);

        File projdir = new File(selectedProject.getProjectDirectory().getPath());

        wiz.putProperty("projdir", projdir);
        AsposeMavenProjectManager asposeMavenProjectManager = AsposeMavenProjectManager.initialize(wiz);
        AsposeWordsJavaAPI.initialize(asposeMavenProjectManager);
        component.read();
    }

    /**
     *
     * @param wiz
     */
    @Override
    public void storeSettings(WizardDescriptor wiz) {

        boolean cancelled = wiz.getValue() != WizardDescriptor.FINISH_OPTION;
        if (!cancelled) {
            if (!storeSettingsCalled) {
                storeSettingsCalled = true;
                createExample();
            } else {
                storeSettingsCalled = false;

            }
        }
    }

    private boolean createExample() {
        String projectPath = component.getSelectedProjectRootPath();
        CustomMutableTreeNode comp = getSelectedNode();
        if (comp == null || !comp.isFolder()) {
            return false;
        }
        try {
           
            String sourceRepositoryExamplePath = comp.getExPath();
             if (sourceRepositoryExamplePath == null) {
                return false;
            }
            String repositorylocation = AsposeWordsJavaAPI.getInstance().getLocalRepositoryPath();
            String repositoryPOM_XML = repositorylocation + File.separator + "Examples" + File.separator + AsposeConstants.MAVEN_POM_XML;

            NodeList examplesNoneAsposeDependencies = AsposeMavenProjectManager.getInstance().getDependenciesFromPOM(repositoryPOM_XML, AsposeConstants.ASPOSE_GROUP_ID);

            AsposeMavenProjectManager.getInstance().addMavenDependenciesInProject(examplesNoneAsposeDependencies);

            NodeList examplesNoneAsposeRepositories = AsposeMavenProjectManager.getInstance().getRepositoriesFromPOM(repositoryPOM_XML, AsposeConstants.ASPOSE_MAVEN_REPOSITORY);

            AsposeMavenProjectManager.getInstance().addMavenRepositoriesInProject(examplesNoneAsposeRepositories);

            String sourceExamplesUtilsPath = repositorylocation + File.separator + AsposeConstants.EXAMPLES_UTIL;
            String destinationExamplesUtilsPath = projectPath + File.separator + sourceExamplesUtilsPath.replace(repositorylocation + File.separator + AsposeConstants.GITHUB_EXAMPLES_SOURCE_LOCATION, AsposeConstants.PROJECT_EXAMPLES_SOURCE_LOCATION);

            String destinationExamplePath = projectPath + File.separator + sourceRepositoryExamplePath.replace(repositorylocation + File.separator + AsposeConstants.GITHUB_EXAMPLES_SOURCE_LOCATION, AsposeConstants.PROJECT_EXAMPLES_SOURCE_LOCATION);

            String destinationResourcePath = projectPath + File.separator + sourceRepositoryExamplePath.replace(repositorylocation + File.separator + AsposeConstants.GITHUB_EXAMPLES_SOURCE_LOCATION, AsposeConstants.PROJECT_EXAMPLES_RESOURCES_LOCATION);

            String sourceRepositoryExampleResourcesPath = sourceRepositoryExamplePath.replace(AsposeConstants.GITHUB_EXAMPLES_SOURCE_LOCATION, AsposeConstants.GITHUB_EXAMPLES_RESOURCES_LOCATION);

            //Copying Example Category
            copyExample(sourceRepositoryExamplePath, destinationExamplePath);

            //Copying Example Resoureces
            copyExample(sourceRepositoryExampleResourcesPath, destinationResourcePath);

            //Copying Utils.java
            AsposeMavenProjectManager.copyDirectory(new File(sourceExamplesUtilsPath + File.separator), new File(destinationExamplesUtilsPath + File.separator));

           

        } catch (Exception ex) {
            return false;
        }
        return true;
    }

    private CustomMutableTreeNode getSelectedNode() {
        return (CustomMutableTreeNode) component.getExamplesTree().getLastSelectedPathComponent();
    }

    private void copyExample(String sourcePath, String destinationPath) {
        try {
            AsposeMavenProjectManager.copyDirectory(sourcePath, destinationPath);
        } catch (IOException ex) {
            Exceptions.printStackTrace(ex);
        }
    }
}
