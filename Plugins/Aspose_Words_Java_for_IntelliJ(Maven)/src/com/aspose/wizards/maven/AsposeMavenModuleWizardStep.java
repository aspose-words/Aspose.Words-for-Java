
package com.aspose.wizards.maven;

import com.aspose.utils.AsposeJavaAPI;
import com.aspose.utils.AsposeMavenProjectManager;
import com.aspose.utils.AsposeConstants;
import com.aspose.utils.AsposeWordsJavaAPI;
import com.aspose.utils.execution.CallBackHandler;
import com.aspose.utils.execution.ModalTaskImpl;
import com.intellij.ide.util.projectWizard.ModuleWizardStep;
import com.intellij.ide.util.projectWizard.WizardContext;
import com.intellij.ide.wizard.CommitStepException;
import com.intellij.openapi.application.ApplicationManager;
import com.intellij.openapi.application.ModalityState;
import com.intellij.openapi.options.ConfigurationException;
import com.intellij.openapi.progress.ProgressManager;
import com.intellij.openapi.project.Project;
import com.intellij.openapi.util.text.StringUtil;
import com.intellij.uiDesigner.core.GridConstraints;
import com.intellij.uiDesigner.core.GridLayoutManager;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ComponentAdapter;
import java.awt.event.ComponentEvent;
import java.util.ResourceBundle;

public class AsposeMavenModuleWizardStep extends ModuleWizardStep {
    private ImageIcon icon = new ImageIcon(getClass().getResource("/resources/long_bannerIntelliJ.png"));


    private final Project myProjectOrNull;
    private final AsposeMavenModuleBuilder myBuilder;
    private final WizardContext myContext;


    private JPanel myMainPanel;


    private JTextField myGroupIdField;

    private JTextField myArtifactIdField;
    private JTextField myVersionField;


    private JPanel myArchetypesPanel;
    private JCheckBox alsoDownloadExampleSourceCheckBox;
    private JTextPane fieldTextPane;
    private JLabel bannerLbl;


    public AsposeMavenModuleWizardStep(Project project, AsposeMavenModuleBuilder builder, WizardContext context, boolean includeArtifacts) {
        myProjectOrNull = project;
        myBuilder = builder;
        myContext = context;
        $$$setupUI$$$();
        initComponents();
        loadSettings();
        bannerLbl.addComponentListener(new ComponentAdapter() {
            @Override
            public void componentResized(ComponentEvent e) {
                int labelwidth = bannerLbl.getWidth();
                int labelheight = bannerLbl.getHeight();
                Image img = icon.getImage();
                bannerLbl.setIcon(new ImageIcon(img.getScaledInstance(labelwidth, labelheight, Image.SCALE_FAST)));
            }
        });
    }

    private void initComponents() {


        ActionListener updatingListener = new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                updateComponents();
            }
        };

    }

    @Override
    public JComponent getPreferredFocusedComponent() {
        return myGroupIdField;
    }


    @Override
    public void onStepLeaving() {
        saveSettings();
    }

    private void loadSettings() {


    }

    private void saveSettings() {

    }
    public JComponent getComponent() {
        return myMainPanel;
    }

    @Override
    public boolean validate() throws ConfigurationException {
        if (StringUtil.isEmptyOrSpaces(myGroupIdField.getText())) {
            throw new ConfigurationException("Please, specify groupId");
        }

        if (StringUtil.isEmptyOrSpaces(myArtifactIdField.getText())) {
            throw new ConfigurationException("Please, specify artifactId");
        }

        if (StringUtil.isEmptyOrSpaces(myVersionField.getText())) {
            throw new ConfigurationException("Please, specify version");
        }

        return true;
    }


    private static void setTestIfEmpty(@NotNull JTextField artifactIdField, @Nullable String text) {
        if (StringUtil.isEmpty(artifactIdField.getText())) {
            artifactIdField.setText(StringUtil.notNullize(text));
        }
    }

    @Override
    public void updateStep() {


        MavenId projectId = myBuilder.getMyProjectId();

        if (projectId == null) {
            setTestIfEmpty(myArtifactIdField, myBuilder.getName());
            setTestIfEmpty(myGroupIdField, myBuilder.getName());
            setTestIfEmpty(myVersionField, "1.0-SNAPSHOT");
        } else {
            setTestIfEmpty(myArtifactIdField, projectId.getArtifactId());
            setTestIfEmpty(myGroupIdField, projectId.getGroupId());
            setTestIfEmpty(myVersionField, projectId.getVersion());
        }


        updateComponents();
    }


    private void updateComponents() {


        myGroupIdField.setEnabled(true);
        myVersionField.setEnabled(true);


    }


    @Override
    public void updateDataModel() {
        myContext.setProjectBuilder(myBuilder);


        myBuilder.setMyProjectId(new MavenId(myGroupIdField.getText(),
                myArtifactIdField.getText(),
                myVersionField.getText()));


        if (myContext.getProjectName() == null) {
            myContext.setProjectName(myBuilder.getMyProjectId().getArtifactId());
        }


    }


    @Override
    public String getHelpId() {
        return "reference.dialogs.new.project.fromScratch.maven";
    }

    @Override
    public void disposeUIResources() {

    }

    @Override
    public void onWizardFinished() throws CommitStepException {

        AsposeMavenProjectManager asposeMavenProjectManager = AsposeMavenProjectManager.initialize(myProjectOrNull);
        AsposeJavaAPI asposeWordsJavaAPI = AsposeWordsJavaAPI.initialize(asposeMavenProjectManager);
        if (alsoDownloadExampleSourceCheckBox.isSelected()) {
            if (!AsposeMavenProjectManager.isInternetConnected()) {

                throw new CommitStepException(AsposeConstants.EXAMPLES_INTERNET_CONNECTION_REQUIRED_MESSAGE);
            }
            CallBackHandler callback = new DownloadExamplesCallback(asposeWordsJavaAPI);

            final ModalTaskImpl modalTaskDownloadExamples = new ModalTaskImpl(myProjectOrNull, callback, ResourceBundle.getBundle("Bundle").getString("AsposeManager.progressExamplesTitle"));
            ApplicationManager.getApplication().invokeAndWait(new Runnable() {
                @Override
                public void run() {
                    ProgressManager.getInstance().run(modalTaskDownloadExamples);
                }
            }, ModalityState.defaultModalityState());

            if (!modalTaskDownloadExamples.isDone()) {

                throw new CommitStepException(AsposeConstants.EXAMPLES_DOWNLOAD_FAIL);
            }
        }

        if (!AsposeMavenProjectManager.isInternetConnected()) {

            throw new CommitStepException(AsposeConstants.MAVEN_INTERNET_CONNECTION_REQUIRED_MESSAGE);
        }

        CallBackHandler callback = new CreateMavenProjectCallback();

        final ModalTaskImpl modalTaskRetrieveArtifact = new ModalTaskImpl(myProjectOrNull, callback, ResourceBundle.getBundle("Bundle").getString("AsposeManager.progressTitle"));
        ApplicationManager.getApplication().invokeAndWait(new Runnable() {
            @Override
            public void run() {
                ProgressManager.getInstance().run(modalTaskRetrieveArtifact);
            }
        }, ModalityState.defaultModalityState());

        if (!modalTaskRetrieveArtifact.isDone()) {

            throw new CommitStepException(AsposeConstants.MAVEN_ARTIFACTS_RETRIEVE_FAIL);
        }



    }


    /**
     * Method generated by IntelliJ IDEA GUI Designer
     * >>> IMPORTANT!! <<<
     * DO NOT edit this method OR call it in your code!
     *
     * @noinspection ALL
     */
    private void $$$setupUI$$$() {
        myMainPanel = new JPanel();
        myMainPanel.setLayout(new GridLayoutManager(3, 3, new Insets(0, 0, 0, 0), -1, -1));
        myArchetypesPanel = new JPanel();
        myArchetypesPanel.setLayout(new BorderLayout(0, 0));
        myArchetypesPanel.setInheritsPopupMenu(true);
        myMainPanel.add(myArchetypesPanel, new GridConstraints(2, 0, 1, 3, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, null, null, null, 0, false));
        alsoDownloadExampleSourceCheckBox = new JCheckBox();
        this.$$$loadButtonText$$$(alsoDownloadExampleSourceCheckBox, ResourceBundle.getBundle("Bundle").getString("AsposeWizardPanel.WizardStep.ExampleChkbox"));
        myArchetypesPanel.add(alsoDownloadExampleSourceCheckBox, BorderLayout.SOUTH);
        fieldTextPane = new JTextPane();
        fieldTextPane.setText("");
        myArchetypesPanel.add(fieldTextPane, BorderLayout.WEST);
        bannerLbl = new JLabel();
        bannerLbl.setAlignmentY(0.0f);
        bannerLbl.setHorizontalAlignment(2);
        bannerLbl.setHorizontalTextPosition(2);
        bannerLbl.setIcon(new ImageIcon(getClass().getResource("/resources/long_bannerIntelliJ.png")));
        bannerLbl.setIconTextGap(0);
        bannerLbl.setText("");
        bannerLbl.setVerticalAlignment(1);
        bannerLbl.setVerticalTextPosition(1);
        myMainPanel.add(bannerLbl, new GridConstraints(0, 0, 1, 2, GridConstraints.ANCHOR_NORTH, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, new Dimension(10, -1), new Dimension(390, -1), new Dimension(66666, -1), 1, false));
        final JPanel panel1 = new JPanel();
        panel1.setLayout(new GridLayoutManager(3, 2, new Insets(10, 10, 10, 10), -1, -1));
        panel1.setAlignmentX(0.0f);
        panel1.setAlignmentY(0.0f);
        panel1.setOpaque(true);
        myMainPanel.add(panel1, new GridConstraints(1, 0, 1, 2, GridConstraints.ANCHOR_NORTH, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_WANT_GROW, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, null, null, null, 0, false));
        panel1.setBorder(BorderFactory.createTitledBorder(null, ResourceBundle.getBundle("Bundle").getString("AsposeWizardPanel.WizardStep.headingMavn"), TitledBorder.LEFT, TitledBorder.DEFAULT_POSITION, new Font(panel1.getFont().getName(), panel1.getFont().getStyle(), panel1.getFont().getSize()), new Color(-16777216)));
        final JLabel label1 = new JLabel();
        label1.setText("GroupId");
        panel1.add(label1, new GridConstraints(0, 0, 1, 1, GridConstraints.ANCHOR_WEST, GridConstraints.FILL_NONE, GridConstraints.SIZEPOLICY_FIXED, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        myGroupIdField = new JTextField();
        panel1.add(myGroupIdField, new GridConstraints(0, 1, 1, 1, GridConstraints.ANCHOR_WEST, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_WANT_GROW, GridConstraints.SIZEPOLICY_FIXED, null, new Dimension(444, 24), null, 0, false));
        final JLabel label2 = new JLabel();
        label2.setText("ArtifactId");
        panel1.add(label2, new GridConstraints(1, 0, 1, 1, GridConstraints.ANCHOR_WEST, GridConstraints.FILL_NONE, GridConstraints.SIZEPOLICY_FIXED, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        myArtifactIdField = new JTextField();
        panel1.add(myArtifactIdField, new GridConstraints(1, 1, 1, 1, GridConstraints.ANCHOR_WEST, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_WANT_GROW, GridConstraints.SIZEPOLICY_FIXED, null, new Dimension(444, 24), null, 0, false));
        final JLabel label3 = new JLabel();
        label3.setText("Version");
        panel1.add(label3, new GridConstraints(2, 0, 1, 1, GridConstraints.ANCHOR_WEST, GridConstraints.FILL_NONE, GridConstraints.SIZEPOLICY_FIXED, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        myVersionField = new JTextField();
        myVersionField.setText("");
        panel1.add(myVersionField, new GridConstraints(2, 1, 1, 1, GridConstraints.ANCHOR_WEST, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_WANT_GROW, GridConstraints.SIZEPOLICY_FIXED, null, new Dimension(444, 24), null, 0, false));
    }

    /**
     * @noinspection ALL
     */
    private void $$$loadButtonText$$$(AbstractButton component, String text) {
        StringBuffer result = new StringBuffer();
        boolean haveMnemonic = false;
        char mnemonic = '\0';
        int mnemonicIndex = -1;
        for (int i = 0; i < text.length(); i++) {
            if (text.charAt(i) == '&') {
                i++;
                if (i == text.length()) break;
                if (!haveMnemonic && text.charAt(i) != '&') {
                    haveMnemonic = true;
                    mnemonic = text.charAt(i);
                    mnemonicIndex = result.length();
                }
            }
            result.append(text.charAt(i));
        }
        component.setText(result.toString());
        if (haveMnemonic) {
            component.setMnemonic(mnemonic);
            component.setDisplayedMnemonicIndex(mnemonicIndex);
        }
    }

    /**
     * @noinspection ALL
     */
    public JComponent $$$getRootComponent$$$() {
        return myMainPanel;
    }
}

