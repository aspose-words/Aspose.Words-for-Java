package com.aspose.wizards.maven;

import com.aspose.utils.AsposeMavenProjectManager;
import com.aspose.utils.execution.CallBackHandler;
import com.intellij.openapi.progress.ProgressIndicator;
import org.jetbrains.annotations.NotNull;

import java.util.ResourceBundle;

/**
 * Created by Adeel Ilyas on 08/19/2015.
 */
public class CreateMavenProjectCallback implements CallBackHandler {

    @Override
    public boolean executeTask(@NotNull ProgressIndicator progressIndicator) {

        progressIndicator.setIndeterminate(true);
        progressIndicator.setText(ResourceBundle.getBundle("Bundle").getString("AsposeManager.projectMessage"));
        AsposeMavenProjectManager comManager = AsposeMavenProjectManager.getInstance();

        return comManager.retrieveAsposeMavenDependencies(progressIndicator);
    }
}
