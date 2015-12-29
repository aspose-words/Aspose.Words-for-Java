/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.aspose.words.maven.examples;

import java.awt.Component;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Set;
import javax.swing.JComponent;
import javax.swing.event.ChangeListener;
import org.netbeans.api.templates.TemplateRegistration;
import org.openide.WizardDescriptor;
import org.openide.util.NbBundle.Messages;

/**
 * @author Adeel Ilyas
 */
@TemplateRegistration(
        folder = "Classes",
        displayName = "#AsposeNewFileWizardIterator_displayName",
        iconBase = "com/aspose/words/maven/Aspose.png",
        position = 10,
        description = "AsposeExampleWizard.html")
@Messages("AsposeNewFileWizardIterator_displayName=Aspose.Words Code Example")
public final class AsposeExampleWizardIterator implements WizardDescriptor.InstantiatingIterator<WizardDescriptor> {

    private int index;

    private WizardDescriptor wizard;
    private List<WizardDescriptor.Panel<WizardDescriptor>> panels;

    private List<WizardDescriptor.Panel<WizardDescriptor>> getPanels() {
        if (panels == null) {
            panels = new ArrayList<>();

            panels.add(new AsposeExampleWizardPanel());
            String[] steps = createSteps();

            for (int i = 0; i < panels.size(); i++) {
                Component c = panels.get(i).getComponent();
                if (steps[i] == null) {
                    // Default step name to component name of panel. Mainly
                    // useful for getting the name of the target chooser to
                    // appear in the list of steps.
                    steps[i] = c.getName();
                }
                if (c instanceof JComponent) { // assume Swing components
                    JComponent jc = (JComponent) c;
                    jc.putClientProperty(WizardDescriptor.PROP_CONTENT_SELECTED_INDEX, i);
                    jc.putClientProperty(WizardDescriptor.PROP_CONTENT_DATA, steps);
                    jc.putClientProperty(WizardDescriptor.PROP_AUTO_WIZARD_STYLE, true);
                    jc.putClientProperty(WizardDescriptor.PROP_CONTENT_DISPLAYED, true);
                    jc.putClientProperty(WizardDescriptor.PROP_CONTENT_NUMBERED, true);
                }
            }
        }
        return panels;
    }

    /**
     *
     * @return
     * @throws IOException
     */
    @Override
    public Set<?> instantiate() throws IOException {
        // TODO return set of FileObject (or DataObject) you have created
        return Collections.emptySet();
    }

    /**
     *
     * @param wizard
     */
    @Override
    public void initialize(WizardDescriptor wizard) {
        this.wizard = wizard;
    }

    /**
     *
     * @param wizard
     */
    @Override
    public void uninitialize(WizardDescriptor wizard) {
        panels = null;
    }

    /**
     *
     * @return
     */
    @Override
    public WizardDescriptor.Panel<WizardDescriptor> current() {
        return getPanels().get(index);
    }

    /**
     *
     * @return
     */
    @Override
    public String name() {
        return index + 1 + ". from " + getPanels().size();
    }

    /**
     *
     * @return
     */
    @Override
    public boolean hasNext() {
        return index < getPanels().size() - 1;
    }

    /**
     *
     * @return
     */
    @Override
    public boolean hasPrevious() {
        return index > 0;
    }

    /**
     *
     */
    @Override
    public void nextPanel() {
        if (!hasNext()) {
            throw new NoSuchElementException();
        }
        index++;
    }

    /**
     *
     */
    @Override
    public void previousPanel() {
        if (!hasPrevious()) {
            throw new NoSuchElementException();
        }
        index--;
    }

    // If nothing unusual changes in the middle of the wizard, simply:

    /**
     *
     * @param l
     */
    @Override
    public void addChangeListener(ChangeListener l) {
    }

    /**
     *
     * @param l
     */
    @Override
    public void removeChangeListener(ChangeListener l) {
    }

    private String[] createSteps() {
        String[] beforeSteps = (String[]) wizard.getProperty("WizardPanel_contentData");
        assert beforeSteps != null : "This wizard may only be used embedded in the template wizard";
        String[] res = new String[(beforeSteps.length - 1) + panels.size()];
        for (int i = 0; i < res.length; i++) {
            if (i < (beforeSteps.length - 1)) {
                res[i] = beforeSteps[i];
            } else {
                res[i] = panels.get(i - beforeSteps.length + 1).getComponent().getName();
            }
        }
        return res;
    }
}
