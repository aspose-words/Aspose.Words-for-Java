
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
package com.aspose.wizards.maven;

import com.aspose.utils.AsposeConstants;
import com.aspose.utils.execution.RunnableHelper;
import com.intellij.ide.util.projectWizard.ModuleBuilder;
import com.intellij.ide.util.projectWizard.ModuleWizardStep;
import com.intellij.ide.util.projectWizard.SettingsStep;
import com.intellij.ide.util.projectWizard.WizardContext;
import com.intellij.openapi.Disposable;
import com.intellij.openapi.module.JavaModuleType;
import com.intellij.openapi.module.ModuleType;
import com.intellij.openapi.module.StdModuleTypes;
import com.intellij.openapi.project.Project;
import com.intellij.openapi.projectRoots.JavaSdkType;
import com.intellij.openapi.projectRoots.SdkTypeId;
import com.intellij.openapi.roots.ModifiableRootModel;
import com.intellij.openapi.roots.ui.configuration.ModulesProvider;
import com.intellij.openapi.util.Disposer;
import com.intellij.openapi.util.io.FileUtil;
import com.intellij.openapi.vfs.LocalFileSystem;
import com.intellij.openapi.vfs.VirtualFile;
import icons.AsposeIcons;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import javax.swing.*;
import java.io.File;
import java.util.ResourceBundle;

/**
 * Author: Adeel Ilyas
 */

public class AsposeMavenModuleBuilder extends ModuleBuilder {

    private Project myProject;
    ResourceBundle bundle = ResourceBundle.getBundle("Bundle");

    @Override
    public String getBuilderId() {
        return getClass().getName();
    }

    @Override
    public Icon getBigIcon() {
        return AsposeIcons.AsposeMedium;
    }

    @Override
    public Icon getNodeIcon() {
        return AsposeIcons.AsposeLogo;
    }


    @Override
    public ModuleWizardStep[] createWizardSteps(@NotNull WizardContext wizardContext, @NotNull ModulesProvider modulesProvider) {
        return new ModuleWizardStep[]{new AsposeMavenModuleWizardStep(getMyProject(), this, wizardContext, !wizardContext.isNewWizard()),

        };
    }

    private VirtualFile createAndGetContentEntry() {
        String path = FileUtil.toSystemIndependentName(getContentEntryPath());
        new File(path).mkdirs();
        return LocalFileSystem.getInstance().refreshAndFindFileByPath(path);
    }

    @Override
    public void setupRootModel(ModifiableRootModel rootModel) throws com.intellij.openapi.options.ConfigurationException {

        final Project project = rootModel.getProject();
        setMyProject(rootModel.getProject());
        final VirtualFile root = createAndGetContentEntry();
        rootModel.addContentEntry(root);

        rootModel.inheritSdk();

        RunnableHelper.runWhenInitialized(getMyProject(), new Runnable() {
            public void run() {

                AsposeMavenModuleBuilderHelper mavenBuilder = new AsposeMavenModuleBuilderHelper(getMyProjectId(), "Create new Maven module", project, root);
                mavenBuilder.configure();

            }
        });

    }

    @Override
    public String getGroupName() {
        return JavaModuleType.JAVA_GROUP;
    }

    public Project getMyProject() {
        return myProject;
    }

    public void setMyProject(Project myProject) {
        this.myProject = myProject;
    }

    @Nullable
    public ModuleWizardStep getCustomOptionsStep(WizardContext context, Disposable parentDisposable) {
        AsposeIntroWizardStep step = new AsposeIntroWizardStep();
        Disposer.register(parentDisposable, step);
        return step;
    }


    public ModuleType getModuleType() {
        return StdModuleTypes.JAVA;
    }

    @Override
    public boolean isSuitableSdkType(SdkTypeId sdkType) {
        return sdkType instanceof JavaSdkType;
    }

    @Nullable
    @Override
    public ModuleWizardStep modifySettingsStep(@NotNull SettingsStep settingsStep) {
        return StdModuleTypes.JAVA.modifySettingsStep(settingsStep, this);
    }


    @Nullable
    protected static String getPathForOutputPathStep() {
        return null;
    }


    public MavenId getMyProjectId() {
        return myProjectId;
    }

    public void setMyProjectId(MavenId myProjectId) {
        this.myProjectId = myProjectId;
    }

    private MavenId myProjectId;


    @Override
    public String getPresentableName() {
        return AsposeConstants.WIZARD_NAME;
    }


    @Override
    public String getDescription() {
        return bundle.getString("AsposeWizardPanel.myMainPanel.description");
    }

}