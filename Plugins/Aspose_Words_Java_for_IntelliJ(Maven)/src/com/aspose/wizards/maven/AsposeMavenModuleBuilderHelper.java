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

import com.aspose.maven.apis.artifacts.Metadata;
import com.aspose.utils.AsposeConstants;
import com.aspose.utils.AsposeMavenProjectManager;
import com.aspose.utils.AsposeMavenUtil;
import com.intellij.codeInsight.actions.ReformatCodeProcessor;
import com.intellij.ide.util.EditorHelper;
import com.intellij.openapi.application.ModalityState;
import com.intellij.openapi.application.Result;
import com.intellij.openapi.command.WriteCommandAction;
import com.intellij.openapi.project.Project;
import com.intellij.openapi.project.ex.ProjectManagerEx;
import com.intellij.openapi.util.io.FileUtil;
import com.intellij.openapi.vfs.LocalFileSystem;
import com.intellij.openapi.vfs.VfsUtil;
import com.intellij.openapi.vfs.VirtualFile;
import com.intellij.psi.PsiFile;
import com.intellij.psi.PsiManager;
import com.sun.xml.internal.messaging.saaj.util.ByteOutputStream;
import org.jetbrains.annotations.NotNull;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

public class AsposeMavenModuleBuilderHelper {
    private final MavenId myProjectId;

    private final String myCommandName;
    private Project project;
    private VirtualFile root;
    private List<String> intelliJMavenFiles = new ArrayList<String>();

    public AsposeMavenModuleBuilderHelper(@NotNull MavenId projectId, String commaneName, Project project, VirtualFile root) {
        myProjectId = projectId;
        this.project = project;
        this.root = root;
        myCommandName = commaneName;
        intelliJMavenFiles.add("untitled.iml");
        intelliJMavenFiles.add("compiler.xml");
    }

    public void configure() {
        PsiFile[] psiFiles = PsiFile.EMPTY_ARRAY;
        final VirtualFile pom = new WriteCommandAction<VirtualFile>(project, myCommandName, psiFiles) {
            @Override
            protected void run(Result<VirtualFile> result) throws Throwable {
                VirtualFile file;
                try {
                    file = root.createChildData(this, AsposeConstants.MAVEN_POM_XML);

                    AsposeMavenUtil.runOrApplyMavenProjectFileTemplate(project, file, myProjectId);
                    result.setResult(file);
                } catch (IOException e) {
                    showError(project, e);
                    return;
                }

                updateProjectPom(project, file);


            }
        }.execute().getResultObject();

        if (pom == null) return;

        try {
            System.out.println("Creating Maven project structure ...");
            VfsUtil.createDirectories(root.getPath() + "/src/main/java");
            VfsUtil.createDirectories(root.getPath() + "/src/main/resources");
            VfsUtil.createDirectories(root.getPath() + "/src/test/java");
        } catch (IOException e) {

        }
        // execute when current dialog is closed (e.g. Project Structure)
        AsposeMavenUtil.invokeLater(project, ModalityState.NON_MODAL, new Runnable() {
            public void run() {
                if (!pom.isValid()) return;
                copyMavenConfigurationFiles(pom);


            }
        });
    }

    private void writeXmlDocumentToVirtualFile(VirtualFile pom, Document pomDocument) throws TransformerConfigurationException, TransformerException, IOException {
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");
        DOMSource source = new DOMSource(pomDocument);

        ByteOutputStream bytes = new ByteOutputStream();

        StreamResult result = new StreamResult(bytes);
        transformer.transform(source, result);
        VfsUtil.saveText(pom, bytes.toString());
    }

    private Document getXmlDocument(String xmlfile) throws ParserConfigurationException, SAXException, IOException {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        Document pomDocument = docBuilder.parse(xmlfile);

        return pomDocument;
    }

    private void addAsposeMavenRepositoryConfiguration(Document pomDocument, Node projectNode) {
// Adding Aspose Cloud Maven Repository configuration
        Element repositories = pomDocument.createElement("repositories");
        projectNode.appendChild(repositories);
        Element repository = pomDocument.createElement("repository");
        repositories.appendChild(repository);
        Element id = pomDocument.createElement("id");
        id.appendChild(pomDocument.createTextNode("AsposeJavaAPI"));
        Element name = pomDocument.createElement("name");
        name.appendChild(pomDocument.createTextNode("Aspose Java API"));
        Element url = pomDocument.createElement("url");
        url.appendChild(pomDocument.createTextNode("http://maven.aspose.com/artifactory/simple/ext-release-local/"));
        repository.appendChild(id);
        repository.appendChild(name);
        repository.appendChild(url);
    }

    // In misc.xml
    private void addMavenConfiguration(final VirtualFile miscxml, String mavenMiscXmlfile) {
        try {
            Document pomDocument = getXmlDocument(mavenMiscXmlfile);

            // Get the root (Project Node) element
            Node projectNode = pomDocument.getFirstChild();
            Element component = pomDocument.createElement("component");
            component.setAttribute("name", "MavenProjectsManager");
            projectNode.appendChild(component);
            Element option = pomDocument.createElement("option");
            option.setAttribute("name", "originalFiles");
            component.appendChild(option);
            Element list = pomDocument.createElement("list");
            option.appendChild(list);
            Element listOption = pomDocument.createElement("option");
            listOption.setAttribute("value", "$PROJECT_DIR$/pom.xml");
            list.appendChild(listOption);

            // Write the content into misc xml file
            writeXmlDocumentToVirtualFile(miscxml, pomDocument);
        } catch (IOException io) {
            io.printStackTrace();
        } catch (ParserConfigurationException pce) {
            pce.printStackTrace();
        } catch (TransformerException tfe) {
            tfe.printStackTrace();
        } catch (SAXException sae) {
            sae.printStackTrace();
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

    private void updateProjectPom(final Project project, final VirtualFile pom) {
        try {
            String mavenPomXmlfile = AsposeMavenUtil.getPOMXmlFile(pom);

            Document pomDocument = getXmlDocument(mavenPomXmlfile);

            // Get the root element
            Node projectNode = pomDocument.getFirstChild();

            //Adding Aspose Cloud Maven Repository configuration setting here
            addAsposeMavenRepositoryConfiguration(pomDocument, projectNode);

            // Adding Dependencies here
            Element dependenciesTag = pomDocument.createElement("dependencies");
            projectNode.appendChild(dependenciesTag);

            for (Metadata dependency : getAsposeProjectMavenDependencies()) {

                addAsposeMavenDependency(pomDocument, dependenciesTag, dependency);

            }

            // Write the content into maven pom xml file

            writeXmlDocumentToVirtualFile(pom, pomDocument);


        } catch (IOException io) {
            io.printStackTrace();
        } catch (ParserConfigurationException pce) {
            pce.printStackTrace();
        } catch (TransformerException tfe) {
            tfe.printStackTrace();
        } catch (SAXException sae) {
            sae.printStackTrace();
        }
    }

    private static PsiFile getPsiFile(Project project, VirtualFile pom) {
        return PsiManager.getInstance(project).findFile(pom);
    }

    public String getAsposeMavenWorkSpace() {
        final String RepositoryResourcesLocation = "https://raw.githubusercontent.com/asposemarketplace/Aspose_Maven_for_JetBrains/master/src/resources/maven/";

        String path = "";
        path = System.getProperty("user.home");
        path = path + File.separator + "aspose" + File.separator + "intellijplugin" + File.separator + "maven" + File.separator;
        File confirmPath = new File(path);
        if (!confirmPath.exists()) {
            new File(path).mkdirs();
            for (String fileToDownload : intelliJMavenFiles) {
                downloadFileFromInternet(RepositoryResourcesLocation + fileToDownload, path + fileToDownload);
            }

        }
        return path;
    }

    public boolean downloadFileFromInternet(String urlStr, String absoluteOutputFile) {
        InputStream input;
        int bufferSize = 4096;

        try {

            URL url = new URL(urlStr);
            input = url.openStream();
            byte[] buffer = new byte[bufferSize];
            File f = new File(absoluteOutputFile);
            OutputStream output = new FileOutputStream(f);


            try {
                int bytesRead;
                while ((bytesRead = input.read(buffer, 0, buffer.length)) >= 0) {
                    output.write(buffer, 0, bytesRead);


                }

                output.flush();
                output.close();

            } finally {
            }
        } catch (Exception ex) {
            return false;
        }
        return true;
    }

    private void copyMavenConfigurationFiles(VirtualFile pom) {
        try {


            String projectPath = project.getBasePath();

            final File workingDir = new File(getAsposeMavenWorkSpace());

            String projectModulefile = projectPath + File.separator + project.getName() + ".iml";
            String projectIdea_compiler_xml = projectPath + File.separator + ".idea" + File.separator + "compiler.xml";

            String projectIdea_misc_xml = projectPath + File.separator + ".idea" + File.separator + "misc.xml";
            VirtualFile vf_projectIdea_misc_xml = LocalFileSystem.getInstance().findFileByPath(projectIdea_misc_xml);

            FileUtil.copy(new File(workingDir, intelliJMavenFiles.get(0)), new File(projectModulefile));
            FileUtil.copy(new File(workingDir, intelliJMavenFiles.get(1)), new File(projectIdea_compiler_xml));

            addMavenConfiguration(vf_projectIdea_misc_xml,projectIdea_misc_xml);

            ProjectManagerEx pm = ProjectManagerEx.getInstanceEx();

            pm.reloadProject(project);

            EditorHelper.openInEditor(getPsiFile(project, pom));

        } catch (IOException e) {
            showError(project, e);
            return;
        } catch (Throwable e) {
        }

    }

    private static void showError(Project project, Throwable e) {
        AsposeMavenUtil.showError(project, "Failed to create a Maven project", e);
    }

    private static void updateFileContents(Project project, final VirtualFile vf, final File f) throws Throwable {
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        InputStream in = null;
        try {

            in = new FileInputStream(f);

            write(in, bytes);

        } finally {

            if (in != null) {
                in.close();
            }
        }
        VfsUtil.saveText(vf, bytes.toString());

        PsiFile psiFile = PsiManager.getInstance(project).findFile(vf);
        if (psiFile != null) {
            new ReformatCodeProcessor(project, psiFile, null, false).run();
        }
    }

    private static void write(InputStream inputStream, OutputStream os) throws IOException {
        byte[] buf = new byte[1024];
        int len;
        while ((len = inputStream.read(buf)) > 0) {
            os.write(buf, 0, len);
        }
    }

    public static List<Metadata> getAsposeProjectMavenDependencies() {
        return asposeProjectMavenDependencies;
    }

    private static List<Metadata> asposeProjectMavenDependencies = new ArrayList<Metadata>();
}
