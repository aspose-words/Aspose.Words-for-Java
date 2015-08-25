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

package com.aspose.utils;

import com.intellij.openapi.progress.ProgressIndicator;
import org.jetbrains.annotations.NotNull;

import javax.swing.*;
import java.io.File;

/**
 * Created by Adeel on 8/13/2015.
 */
public abstract class AsposeJavaAPI {
    public abstract String get_name();

    public abstract String get_mavenRepositoryURL();

    public abstract String get_remoteExamplesRepository();

    public boolean isExamplesNotAvailable() {
        return examplesNotAvailable;
    }

    public boolean examplesNotAvailable;

    public boolean isExamplesDefinitionAvailable() {
        return examplesDefinitionAvailable;
    }

    public boolean examplesDefinitionAvailable;

    public AsposeMavenProjectManager asposeMavenProjectManager;

    public void checkAndUpdateRepo(ProgressIndicator p) {

        if (null == get_remoteExamplesRepository()) {
            AsposeMavenProjectManager.showMessage(AsposeConstants.EXAMPLES_NOT_AVAILABLE_TITLE, get_name() + " - " + AsposeConstants.EXAMPLES_NOT_AVAILABLE_MESSAGE, JOptionPane.CLOSED_OPTION, JOptionPane.INFORMATION_MESSAGE);
            examplesNotAvailable = true;
            examplesDefinitionAvailable = false;
            return;
        } else {
            examplesNotAvailable = false;
        }

        if (isExamplesDefinitionsPresent()) {
            try {
                examplesDefinitionAvailable = true;
                syncRepository(p);
                p.setFraction(0.30);
            } catch (Exception e) {
            }
        } else {
            updateRepository(p);
            if (isExamplesDefinitionsPresent()) {
                examplesDefinitionAvailable = true;

            }


        }
        p.setFraction(0.50);
    }
    public boolean downloadExamples(@NotNull ProgressIndicator progressIndicator) {
        try {
            checkAndUpdateRepo(progressIndicator);
        } catch (Exception rex) {
            return false;
        }

        return true;

    }


    public void updateRepository(ProgressIndicator p)
    {
        AsposeMavenProjectManager.checkAndCreateFolder(getLocalRepositoryPath());

        try {

            GitHelper.updateRepository(getLocalRepositoryPath(), get_remoteExamplesRepository());
            p.setFraction(1);

        } catch (Exception e) {
        }
    }

    public void syncRepository(ProgressIndicator p)
    {   try {

            GitHelper.syncRepository(getLocalRepositoryPath(), get_remoteExamplesRepository());
            p.setFraction(1);

        } catch (Exception e) {
        }
    }


    /**
     *
     * @return boolean
     */
    public boolean isExamplesDefinitionsPresent()
    {
        return new File(getLocalRepositoryPath()).exists();
    }

    /**
     *
     * @return String
     */
    public String getLocalRepositoryPath()
    {
        return asposeMavenProjectManager.getAsposeHomePath() +  "GitConsRepos" + File.separator + get_name();
    }
}
