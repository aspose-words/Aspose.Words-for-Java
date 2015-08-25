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

package com.aspose.examples;

import com.aspose.utils.*;
import com.intellij.openapi.diagnostic.Logger;
import com.intellij.openapi.ui.DialogWrapper;

import javax.swing.*;
import java.io.File;
import java.io.IOException;

/**
 * Created by Adeel Ilyas on 8/17/2015.
 */

public class AsposeExampleDialog extends DialogWrapper {
    private static final Logger LOG = Logger.getInstance("#com.aspose.examples.AsposeExampleDialog");

    private final String myDescription;
    private AsposeExamplePanel component;

    public AsposeExampleDialog(final String title, String description) {
        super(false);
        myDescription = description;
        setTitle(title);
        init();
        setOKButtonText("Create");
        setOKActionEnabled(false);
    }

    public void updateControls(boolean selection) {
        setOKActionEnabled(selection);
    }

    @Override
    protected void doOKAction() {
        super.doOKAction();
        createExample();
        AsposeMavenProjectManager.getInstance().getProjectHandle().getProjectFile().getFileSystem().refresh(false);
        AsposeMavenProjectManager.getInstance().getProjectHandle().getBaseDir().getFileSystem().refresh(false);
    }

    @Override
    public JComponent getPreferredFocusedComponent() {
        AsposeConstants.println("AsposeExamplePanel getComponent(): is called ...");
        if (component == null) {
            component = new AsposeExamplePanel(this);
        }

        return component.getComponentSelection();
    }

    @Override
    protected JComponent createCenterPanel() {
        AsposeConstants.println("AsposeExamplePanel getComponent(): is called ...");
        if (component == null) {
            component = new AsposeExamplePanel(this);
        }
        return component;
    }

    @Override
    protected String getDimensionServiceKey() {
        return "#com.aspose.examples.AsposeExampleDialog";
    }

    //=========================================================================


    //=========================================================================
    private boolean createExample() {
        String projectPath = component.getSelectedProjectRootPath();
        CustomMutableTreeNode comp = getSelectedNode();
        if (comp == null || !comp.isFolder()) {
            return false;
        }
        try {

            String sourceRepositoryExamplePath = comp.getExPath();

            String repositorylocation = AsposeWordsJavaAPI.getInstance().getLocalRepositoryPath();

            String sourceExamplesUtilsPath= repositorylocation+File.separator+AsposeConstants.REPOSITORY_UTIL;
            String destinationExamplesUtilsPath = projectPath + File.separator + sourceExamplesUtilsPath.replace(repositorylocation + File.separator + AsposeConstants.SOURCE_API_EXAMPLES_LOCATION, AsposeConstants.DESTINATION_API_EXAMPLES_LOCATION);

            String destinationExamplePath = projectPath + File.separator + sourceRepositoryExamplePath.replace(repositorylocation + File.separator + AsposeConstants.SOURCE_API_EXAMPLES_LOCATION, AsposeConstants.DESTINATION_API_EXAMPLES_LOCATION);

            String destinationResourcePath = projectPath + File.separator + sourceRepositoryExamplePath.replace(repositorylocation + File.separator + AsposeConstants.SOURCE_API_EXAMPLES_LOCATION, AsposeConstants.DESTINATION_API_EXAMPLES_RESOURCES_LOCATION);


            String sourceRepositoryExampleResourcesPath = sourceRepositoryExamplePath.replace(AsposeConstants.SOURCE_API_EXAMPLES_LOCATION, AsposeConstants.SOURCE_API_EXAMPLES_RESOURCES_LOCATION);

            //Copying Example Category
            copyExample(sourceRepositoryExamplePath, destinationExamplePath);

            //Copying Example Resoureces
            copyExample(sourceRepositoryExampleResourcesPath, destinationResourcePath);


            //Copying Utils.java
            AsposeMavenProjectManager.copyDirectory(new File(sourceExamplesUtilsPath + File.separator), new File(destinationExamplesUtilsPath + File.separator));

            if (sourceRepositoryExamplePath == null || comp == null) {
                return false;
            }

        } catch (Exception ex) {
            return false;
        }
        return true;
    }

    //=========================================================================
    private CustomMutableTreeNode getSelectedNode() {
        return (CustomMutableTreeNode) component.getExamplesTree().getLastSelectedPathComponent();
    }

    //=========================================================================
    private void copyExample(String sourcePath, String destinationPath) {
        try {
            AsposeMavenProjectManager.copyDirectory(sourcePath, destinationPath);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}
