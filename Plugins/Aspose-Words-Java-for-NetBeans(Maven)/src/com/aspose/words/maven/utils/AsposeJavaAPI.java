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
package com.aspose.words.maven.utils;

import javax.swing.*;
import java.io.File;
import org.netbeans.api.progress.aggregate.ProgressContributor;
import org.openide.util.Exceptions;

/**
 * Created by Adeel on 12/16/2015.
 */
public abstract class AsposeJavaAPI {

    /**
     *
     * @return
     */
    public abstract String get_name();

    /**
     *
     * @return
     */
    public abstract String get_mavenRepositoryURL();

    /**
     *
     * @return
     */
    public abstract String get_remoteExamplesRepository();

    /**
     *
     * @return
     */
    public boolean isExamplesNotAvailable() {
        return examplesNotAvailable;
    }

    /**
     *
     */
    public boolean examplesNotAvailable;

    /**
     *
     * @return
     */
    public boolean isExamplesDefinitionAvailable() {
        return examplesDefinitionAvailable;
    }

    /**
     *
     */
    public boolean examplesDefinitionAvailable;

    /**
     *
     */
    public AsposeMavenProjectManager asposeMavenProjectManager;

    /**
     *
     * @param p
     */
    public void checkAndUpdateRepo(ProgressContributor p) {

        if (null == get_remoteExamplesRepository()) {
            AsposeMavenProjectManager.showMessage(AsposeConstants.EXAMPLES_NOT_AVAILABLE_TITLE, get_name() + " - " + AsposeConstants.EXAMPLES_NOT_AVAILABLE_MSG, JOptionPane.CLOSED_OPTION, JOptionPane.INFORMATION_MESSAGE);
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
                p.progress(60);
            } catch (Exception e) {
            }
        } else {
            updateRepository(p);
            if (isExamplesDefinitionsPresent()) {
                examplesDefinitionAvailable = true;

            }

        }
        p.progress(70);
    }

    /**
     *
     * @param p
     * @return
     */
    public boolean downloadExamples(ProgressContributor p) {
        try {
            checkAndUpdateRepo(p);
        } catch (Exception rex) {
            Exceptions.printStackTrace(rex);
            return false;
        }

        return true;

    }

    /**
     *
     * @param p
     */
    public void updateRepository(ProgressContributor p) {
        AsposeMavenProjectManager.checkAndCreateFolder(getLocalRepositoryPath());

        try {

            GitHelper.updateRepository(getLocalRepositoryPath(), get_remoteExamplesRepository());
            p.progress(55);

        } catch (Exception e) {
            Exceptions.printStackTrace(e);
        }
    }

    /**
     *
     * @param p
     */
    public void syncRepository(ProgressContributor p) {
        try {

            GitHelper.syncRepository(getLocalRepositoryPath(), get_remoteExamplesRepository());
            p.progress(55);

        } catch (Exception e) {
            Exceptions.printStackTrace(e);
        }
    }

    /**
     *
     * @return boolean
     */
    public boolean isExamplesDefinitionsPresent() {
        return new File(getLocalRepositoryPath()).exists();
    }

    /**
     *
     * @return String
     */
    public String getLocalRepositoryPath() {
        return asposeMavenProjectManager.getAsposeHomePath() + "GitConsRepos" + File.separator + get_name();
    }
}
