package com.aspose.words.maven;

import com.aspose.words.maven.artifacts.Metadata;
import com.aspose.words.maven.utils.AsposeConstants;
import com.aspose.words.maven.utils.AsposeJavaAPI;
import com.aspose.words.maven.utils.AsposeMavenProjectManager;
import static com.aspose.words.maven.utils.AsposeMavenProjectManager.getAsposeProjectMavenDependencies;
import com.aspose.words.maven.utils.AsposeWordsJavaAPI;
import com.aspose.words.maven.utils.TasksExecutor;
import java.awt.Component;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import javax.swing.JComponent;
import javax.swing.event.ChangeListener;
import org.netbeans.api.progress.ProgressHandle;
import org.netbeans.api.project.ProjectManager;
import org.netbeans.api.templates.TemplateRegistration;
import org.netbeans.spi.project.ui.support.ProjectChooser;
import org.netbeans.spi.project.ui.templates.support.Templates;
import org.openide.WizardDescriptor;
import org.openide.filesystems.FileObject;
import org.openide.filesystems.FileUtil;
import org.openide.util.Exceptions;
import org.openide.util.NbBundle;
import org.openide.util.NbBundle.Messages;
import org.openide.xml.XMLUtil;
import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

/**
 * @author Adeel Ilyas
 */
@TemplateRegistration(
        folder = "Project/Maven2",
        displayName = "#Aspose_displayName",
        description = "AsposeWordsMavenDescription.html",
        iconBase = "com/aspose/words/maven/Aspose.png",
        position = 1,
        content = "AsposeMavenProject.zip")
@Messages("Aspose_displayName=Aspose.Words Maven Project")

public class AsposeMavenProjectWizardIterator implements WizardDescriptor.ProgressInstantiatingIterator<WizardDescriptor> {

    private int index;
    private WizardDescriptor.Panel[] panels;
    private WizardDescriptor wiz;
    List<String> list = new ArrayList<>();

    /**
     *
     */
    public AsposeMavenProjectWizardIterator() {
    }

    /**
     *
     * @return
     */
    public static AsposeMavenProjectWizardIterator createIterator() {
        return new AsposeMavenProjectWizardIterator();
    }

    private WizardDescriptor.Panel[] createPanels() {

        return new WizardDescriptor.Panel[]{
            new AsposeMavenBasicWizardPanel()
        };
    }

     /**
     *
     * @return
     */
    private String[] createSteps() {
        return new String[]{
            NbBundle.getMessage(AsposeMavenProjectWizardIterator.class, "LBL_CreateProjectStep"),
        };
    }

    /**
     *
     * @return
     * @throws IOException
     */
    @Override
    public Set<?> instantiate() throws IOException {
        throw new AssertionError("instantiate(ProgressHandle) " //NOI18N
                + "should have been called"); //NOI18N
    }

    /**
     *
     * @param ph
     * @return
     * @throws IOException
     */
    @Override
    public Set instantiate(ProgressHandle ph) throws IOException {
        ph.start();
        ph.switchToIndeterminate();
        ph.progress("Processing...");
        final AsposeMavenProjectManager asposeMavenProjectManager = AsposeMavenProjectManager.initialize(wiz);
        final AsposeJavaAPI asposeWordsJavaAPI = AsposeWordsJavaAPI.initialize(asposeMavenProjectManager);

        boolean isDownloadExamplesSelected = (boolean) wiz.getProperty("downloadExamples");

        // Downloading Aspose.Words Java (mvn based) examples...
        if (isDownloadExamplesSelected) {
            TasksExecutor tasksExecutionDownloadExamples = new TasksExecutor(NbBundle.getMessage(AsposeMavenProjectWizardIterator.class, "AsposeManager.progressExamplesTitle"));
            // Downloading Aspose API mvn based examples
            tasksExecutionDownloadExamples.addNewTask(asposeMavenProjectManager.createDownloadExamplesTask(asposeWordsJavaAPI));
            // Execute the tasks
            tasksExecutionDownloadExamples.processTasks();
        }
        TasksExecutor tasksExecutionRetrieve = new TasksExecutor(NbBundle.getMessage(AsposeMavenProjectWizardIterator.class, "AsposeManager.progressTitle"));

        // Retrieving Aspose.Words Java Maven artifact...
        tasksExecutionRetrieve.addNewTask(asposeMavenProjectManager.retrieveAsposeAPIMavenTask(asposeWordsJavaAPI));

        // Execute the tasks
        tasksExecutionRetrieve.processTasks();

        // Creating Maven project
        ph.progress(NbBundle.getMessage(AsposeMavenProjectWizardIterator.class, "AsposeManager.projectMessage"));

        Set<FileObject> resultSet = new LinkedHashSet<>();

        File projectDir = FileUtil.normalizeFile((File) wiz.getProperty("projdir"));
        projectDir.mkdirs();

        FileObject template = Templates.getTemplate(wiz);
        FileObject projectRoot = FileUtil.toFileObject(projectDir);
        createAsposeMavenProject(template.getInputStream(), projectRoot);

        createStartupPackage(projectRoot);

        resultSet.add(projectRoot);
        // Look for nested projects to open as well:
        Enumeration<? extends FileObject> e = projectRoot.getFolders(true);
        while (e.hasMoreElements()) {
            FileObject subfolder = e.nextElement();
            if (ProjectManager.getDefault().isProject(subfolder)) {
                resultSet.add(subfolder);
            }
        }

        File parent = projectDir.getParentFile();
        if (parent != null && parent.exists()) {
            ProjectChooser.setProjectsFolder(parent);
        }
        ph.finish();
        return resultSet;
    }

    /**
     *
     * @param wiz
     */
    @Override
    public void initialize(WizardDescriptor wiz) {
        this.wiz = wiz;
        index = 0;
        panels = createPanels();
        // Make sure list of steps is accurate.
        String[] steps = createSteps();
        for (int i = 0; i < panels.length; i++) {
            Component c = panels[i].getComponent();
            if (steps[i] == null) {
                // Default step name to component name of panel.
                // Mainly useful for getting the name of the target
                // chooser to appear in the list of steps.
                steps[i] = c.getName();
            }
            if (c instanceof JComponent) { // assume Swing components
                JComponent jc = (JComponent) c;
                // Step #.
                // TODO if using org.openide.dialogs >= 7.8, can use WizardDescriptor.PROP_*:
                jc.putClientProperty("WizardPanel_contentSelectedIndex", i);
                // Step name (actually the whole list for reference).
                jc.putClientProperty("WizardPanel_contentData", steps);
            }
        }
    }

    /**
     *
     * @param wiz
     */
    @Override
    public void uninitialize(WizardDescriptor wiz) {
        this.wiz.putProperty("projdir", null);
        this.wiz.putProperty("name", null);
        this.wiz = null;
        panels = null;
    }

    /**
     *
     * @return
     */
    @Override
    public String name() {
        return MessageFormat.format("{0} of {1}",
                new Object[]{
                    index + 1, panels.length
                });
    }

    /**
     *
     * @return
     */
    @Override
    public boolean hasNext() {
        return index < panels.length - 1;
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

    /**
     *
     * @return
     */
    @Override
    public WizardDescriptor.Panel current() {
        return panels[index];
    }

    /**
     *
     * @param l
     */
    @Override
    public final void addChangeListener(ChangeListener l) {
    }

    /**
     *
     * @param l
     */
    @Override
    public final void removeChangeListener(ChangeListener l) {
    }

    private void createAsposeMavenProject(InputStream source, FileObject projectRoot) throws IOException {
        try {
            ZipInputStream str = new ZipInputStream(source);
            ZipEntry entry;
            while ((entry = str.getNextEntry()) != null) {
                if (entry.isDirectory()) {
                    FileUtil.createFolder(projectRoot, entry.getName());
                } else {
                    FileObject fo = FileUtil.createData(projectRoot, entry.getName());
                    if (AsposeConstants.MAVEN_POM_XML.equals(entry.getName())) {
                        /*
                         Special handling for maven pom.xml:
                         1. Defining / creating project artifacts
                         2. Adding latest Aspose.Words Maven Dependency reference into pom.xml
                         */
                        configureProjectMavenPOM(fo, str);
                    } else {
                        writeFile(str, fo);
                    }
                }
            }
        } finally {
            source.close();
        }
    }

    private void createStartupPackage(FileObject projectRoot) throws IOException {
        String startupPackage = wiz.getProperty("package").toString().replace(".", File.separator);

        FileUtil.createFolder(projectRoot, "src" + File.separator + "main" + File.separator + "java" + File.separator + startupPackage);
    }

    private static void writeFile(ZipInputStream str, FileObject fo) throws IOException {
        try (OutputStream out = fo.getOutputStream()) {
            FileUtil.copy(str, out);
        }
    }

    private void configureProjectMavenPOM(FileObject fo, ZipInputStream str) throws IOException {

        String groupId = (String) wiz.getProperty("groupId");
        String artifactId = (String) wiz.getProperty("artifactId");
        String version = (String) wiz.getProperty("version");

        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            FileUtil.copy(str, baos);
            Document doc = XMLUtil.parse(new InputSource(new ByteArrayInputStream(baos.toByteArray())), false, false, null, null);
            Element root = doc.getDocumentElement();
            Node node = root.getElementsByTagName("groupId").item(0);
            node.setTextContent(groupId);

            node = root.getElementsByTagName("artifactId").item(0);
            node.setTextContent(artifactId);

            node = root.getElementsByTagName("version").item(0);
            node.setTextContent(version);

            updateProjectPom(doc);

            try (OutputStream out = fo.getOutputStream()) {
                XMLUtil.write(doc, out, "UTF-8");
            }
        } catch (IOException | SAXException | DOMException ex) {
            Exceptions.printStackTrace(ex);
            writeFile(str, fo);
        }

    }

    private void updateProjectPom(Document pomDocument) {

        // Get the root element
        Node projectNode = pomDocument.getFirstChild();

        // Adding Dependencies here
        Element dependenciesTag = pomDocument.createElement("dependencies");
        projectNode.appendChild(dependenciesTag);

        for (Metadata dependency : getAsposeProjectMavenDependencies()) {

            addAsposeMavenDependency(pomDocument, dependenciesTag, dependency);

        }

    }

    private void addAsposeMavenDependency(Document doc, Element dependenciesTag, Metadata dependency) {
        Element dependencyTag = doc.createElement("dependency");
        dependenciesTag.appendChild(dependencyTag);

        Element groupIdTag = doc.createElement("groupId");
        groupIdTag.appendChild(doc.createTextNode(dependency.getGroupId()));
        dependencyTag.appendChild(groupIdTag);

        Element artifactId = doc.createElement("artifactId");
        artifactId.appendChild(doc.createTextNode(dependency.getArtifactId()));
        dependencyTag.appendChild(artifactId);
        Element version = doc.createElement("version");
        version.appendChild(doc.createTextNode(dependency.getVersioning().getLatest()));
        dependencyTag.appendChild(version);
        if (dependency.getClassifier() != null) {
            Element classifer = doc.createElement("classifier");
            classifer.appendChild(doc.createTextNode(dependency.getClassifier()));
            dependencyTag.appendChild(classifer);
        }
    }
}
