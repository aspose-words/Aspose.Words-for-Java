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

import com.aspose.words.maven.AsposeMavenProjectWizardIterator;
import com.aspose.words.maven.artifacts.Metadata;
import com.aspose.words.maven.artifacts.ObjectFactory;
import com.aspose.words.maven.examples.AsposeExamplePanel;
import com.aspose.words.maven.examples.CustomMutableTreeNode;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;
import java.util.List;
import javax.swing.*;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.Unmarshaller;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.stream.StreamSource;
import javax.xml.xpath.*;
import java.io.*;
import java.net.*;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.Queue;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreePath;
import javax.xml.bind.JAXBException;
import org.openide.WizardDescriptor;
import org.openide.awt.StatusDisplayer;
import org.openide.filesystems.FileObject;
import org.openide.filesystems.FileUtil;
import org.openide.util.Exceptions;
import org.openide.util.NbBundle;
import org.openide.xml.XMLUtil;
import org.w3c.dom.Node;

/*
* @author Adeel Ilyas <adeel.ilyas@aspose.com>
* Date: 12/21/2015
*
 */

/**
 *
 * @author Adeel
 */

public class AsposeMavenProjectManager {

    private boolean examplesNotAvailable;
    private File projectDir = null;

    /**
     *
     * @return
     */
    public File getProjectDir() {
        return projectDir;
    }
    private boolean examplesDefinitionAvailable;

    /**
     *
     * @param Url
     * @return
     * @throws IOException
     */
    public String readURLContents(String Url) throws IOException {
        URL url = new URL(Url);
        URLConnection con = url.openConnection();
        InputStream in = con.getInputStream();
        String encoding = con.getContentEncoding();
        encoding = encoding == null ? "UTF-8" : encoding;
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buf = new byte[8192];
        int len = 0;
        while ((len = in.read(buf)) != -1) {
            baos.write(buf, 0, len);
        }
        String body = new String(baos.toByteArray(), encoding);
        return body;
    }

    /**
     *
     * @param productMavenRepositoryUrl
     * @return
     */
    public Metadata getProductMavenDependency(String productMavenRepositoryUrl) {
        final String mavenMetaDataFileName = "maven-metadata.xml";
        Metadata data = null;

        try {
            String productMavenInfo;
            productMavenInfo = readURLContents(productMavenRepositoryUrl + mavenMetaDataFileName);
            JAXBContext jaxbContext = JAXBContext.newInstance(ObjectFactory.class);
            Unmarshaller unmarshaller;
            unmarshaller = jaxbContext.createUnmarshaller();

            data = (Metadata) unmarshaller.unmarshal(new StreamSource(new StringReader(productMavenInfo)));

            String remoteArtifactFile = productMavenRepositoryUrl + data.getVersioning().getLatest() + "/" + data.getArtifactId() + "-" + data.getVersioning().getLatest();

            if (!remoteFileExists(remoteArtifactFile + ".jar")) {
                AsposeConstants.println("Not Exists");
                data.setClassifier(getResolveSupportedJDK(remoteArtifactFile));
            } else {
                AsposeConstants.println("Exists");
            }
        } catch (IOException | JAXBException ex) {
            Exceptions.printStackTrace(ex);
            data = null;
        }
        return data;
    }

    /**
     *
     * @param ProductURL
     * @return
     */
    public String getResolveSupportedJDK(String ProductURL) {
        String supportedJDKs[] = {"jdk17", "jdk16", "jdk15", "jdk14", "jdk18"};
        String classifier = null;
        for (String jdkCheck : supportedJDKs) {
            if (remoteFileExists(ProductURL + "-" + jdkCheck + ".jar")) {
                AsposeConstants.println("Exists");
                classifier = jdkCheck;
                break;
            } else {
                AsposeConstants.println("Not Exists");
            }
        }
        return classifier;
    }

    /**
     *
     * @param URLName
     * @return
     */
    public boolean remoteFileExists(String URLName) {
        try {
            HttpURLConnection.setFollowRedirects(false);
            // note : you may also need
            //        HttpURLConnection.setInstanceFollowRedirects(false)
            HttpURLConnection con
                    = (HttpURLConnection) new URL(URLName).openConnection();
            con.setRequestMethod("HEAD");
            return (con.getResponseCode() == HttpURLConnection.HTTP_OK);
        } catch (Exception e) {
            Exceptions.printStackTrace(e);
            return false;
        }
    }

    /**
     *
     * @param asposeAPI
     * @return
     */
    public AbstractTask retrieveAsposeAPIMavenTask(final AsposeJavaAPI asposeAPI) {
        return new AbstractTask(NbBundle.getMessage(AsposeMavenProjectWizardIterator.class, "AsposeManager.progressTitle")) {
            @Override
            public void run() {
                String progressMsg = NbBundle.getMessage(AsposeMavenProjectWizardIterator.class, "AsposeManager.progressMessage");

                p.progress(progressMsg);
                StatusDisplayer.getDefault().setStatusText(progressMsg);

                p.start(100);
                p.progress(50);
                retrieveAsposeMavenDependencies();
                StatusDisplayer.getDefault().setStatusText(progressMsg);
                p.progress(100);
                p.finish();
            }
        };
    }

    /**
     *
     * @param asposeAPI
     * @return
     */
    public AbstractTask createDownloadExamplesTask(final AsposeJavaAPI asposeAPI) {
        return new AbstractTask(NbBundle.getMessage(AsposeMavenProjectWizardIterator.class, "AsposeManager.progressExamplesTitle")) {
            @Override
            public void run() {
                String downloadExamplesMessage = NbBundle.getMessage(AsposeMavenProjectWizardIterator.class, "AsposeManager.downloadExamplesMessage");

                p.progress(downloadExamplesMessage);
                StatusDisplayer.getDefault().setStatusText(downloadExamplesMessage);
                p.start(100);
                p.progress(50);
                asposeAPI.downloadExamples(p);
                p.progress(downloadExamplesMessage);
                p.progress(100);
                p.finish();
            }
        };
    }

    /**
     *
     * @param asposeAPI
     * @param panel
     * @return
     */
    public Runnable populateExamplesTask(final AsposeJavaAPI asposeAPI, final AsposeExamplePanel panel) {

        return new Runnable() {
            @Override
            public void run() {
                final CustomMutableTreeNode top = new CustomMutableTreeNode("");
                DefaultTreeModel model = (DefaultTreeModel) panel.getExamplesTree().getModel();
                model.setRoot(top);
                model.reload(top);
                AsposeJavaAPI component = AsposeWordsJavaAPI.getInstance();
                if (component.isExamplesDefinitionAvailable()) {
                    populateExamplesTree(component, top, panel);
                }
                top.setTopTreeNodeText(AsposeConstants.API_NAME);
                model.setRoot(top);
                model.reload(top);
                panel.getExamplesTree().expandPath(new TreePath(top.getPath()));
            }
        };

    }

    /**
     *
     * @return
     */
    public boolean retrieveAsposeMavenDependencies() {
        try {
            getAsposeProjectMavenDependencies().clear();
            AsposeJavaAPI component = AsposeWordsJavaAPI.getInstance();
            Metadata productMavenDependency = getProductMavenDependency(component.get_mavenRepositoryURL());
            if (productMavenDependency != null) {
                getAsposeProjectMavenDependencies().add(productMavenDependency);
            }

        } catch (Exception rex) {
            Exceptions.printStackTrace(rex);
            return false;
        }
        return !getAsposeProjectMavenDependencies().isEmpty();
    }

    /**
     *
     * @return
     */
    public static boolean isInternetConnected() {
        try {
            InetAddress address = InetAddress.getByName(AsposeConstants.INTERNET_CONNNECTIVITY_PING_URL);
            if (address == null) {
                return false;
            }
        } catch (UnknownHostException e) {
            Exceptions.printStackTrace(e);
            return false;
        }

        return true;
    }

    /**
     *
     * @param title
     * @param message
     * @param buttons
     * @param icon
     * @return
     */
    public static int showMessage(String title, String message, int buttons, int icon) {
        int result = JOptionPane.showConfirmDialog(null, message, title, buttons, icon);
        return result;
    }

    private Document getXmlDocument(String mavenPomXmlfile) throws ParserConfigurationException, SAXException, IOException {
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        Document pomDocument = docBuilder.parse(mavenPomXmlfile);

        return pomDocument;
    }

    /**
     *
     * @param dependencyName
     * @return
     */
    public String getDependencyVersionFromPOM(String dependencyName) {
        try {
            String mavenPomXmlfile = projectDir.getPath() + File.separator + AsposeConstants.MAVEN_POM_XML;

            Document pomDocument = getXmlDocument(mavenPomXmlfile);

            XPathFactory xPathfactory = XPathFactory.newInstance();
            XPath xpath = xPathfactory.newXPath();
            String expression = "//version[ancestor::dependency/artifactId[text()='" + dependencyName + "']]";
            XPathExpression xPathExpr = xpath.compile(expression);
            NodeList nl = (NodeList) xPathExpr.evaluate(pomDocument, XPathConstants.NODESET);

            if (nl != null && nl.getLength() > 0) {
                return nl.item(0).getTextContent();
            }
        } catch (IOException | ParserConfigurationException | SAXException | XPathExpressionException e) {
            Exceptions.printStackTrace(e);
        }
        return null;
    }

    /**
     *
     * @return
     */
    public String getAsposeHomePath() {

        return System.getProperty("user.home") + File.separator + "aspose" + File.separator;

    }

    /**
     *
     * @param sourceLocation
     * @param targetLocation
     * @throws IOException
     */
    public static void copyDirectory(String sourceLocation, String targetLocation) throws IOException {

        checkAndCreateFolder(targetLocation);
        copyDirectory(new File(sourceLocation + File.separator), new File(targetLocation + File.separator));
    }

    /**
     *
     * @param sourceLocation
     * @param targetLocation
     * @throws IOException
     */
    public static void copyDirectory(File sourceLocation, File targetLocation) throws IOException {
        if (sourceLocation.isDirectory()) {
            if (!targetLocation.exists()) {
                targetLocation.mkdir();
            }

            String[] children = sourceLocation.list();
            for (String children1 : children) {
                copyDirectory(new File(sourceLocation, children1), new File(targetLocation, children1));
            }
        } else {

            OutputStream out;
            try (InputStream in = new FileInputStream(sourceLocation)) {
                out = new FileOutputStream(targetLocation);
                // Copy the bits from instream to outstream
                byte[] buf = new byte[1024];
                int len;
                while ((len = in.read(buf)) > 0) {
                    out.write(buf, 0, len);
                }
            }
            out.close();
        }
    }

    /**
     *
     * @param folderPath
     */
    public static void checkAndCreateFolder(String folderPath) {
        File folder = new File(folderPath);
        if (!folder.exists()) {
            folder.mkdirs();
        }
    }
    // Singleton instance
    private static AsposeMavenProjectManager asposeMavenProjectManager = new AsposeMavenProjectManager();

    /**
     *
     * @return
     */
    public static AsposeMavenProjectManager getInstance() {
        return asposeMavenProjectManager;
    }

    /**
     *
     * @param wiz
     * @return
     */
    public static AsposeMavenProjectManager initialize(WizardDescriptor wiz) {
        asposeMavenProjectManager = new AsposeMavenProjectManager();
        asposeMavenProjectManager.projectDir = FileUtil.normalizeFile((File) wiz.getProperty("projdir"));
        return asposeMavenProjectManager;
    }

    private AsposeMavenProjectManager() {
    }

    /**
     *
     * @return
     */
    public static List<Metadata> getAsposeProjectMavenDependencies() {
        return asposeProjectMavenDependencies;
    }

    /**
     *
     */
    public static void clearAsposeProjectMavenDependencies() {
        asposeProjectMavenDependencies.clear();
    }

    private static final List<Metadata> asposeProjectMavenDependencies = new ArrayList<Metadata>();

    /**
     *
     * @param addTheseDependencies
     */
    public void addMavenDependenciesInProject(NodeList addTheseDependencies) {

        String mavenPomXmlfile = projectDir.getPath() + File.separator + AsposeConstants.MAVEN_POM_XML;

        try {
            Document pomDocument = getXmlDocument(mavenPomXmlfile);

            Node dependenciesNode = pomDocument.getElementsByTagName("dependencies").item(0);

            if (addTheseDependencies != null && addTheseDependencies.getLength() > 0) {
                for (int n = 0; n < addTheseDependencies.getLength(); n++) {
                    String artifactId = addTheseDependencies.item(n).getFirstChild().getNextSibling().getNextSibling().getNextSibling().getFirstChild().getNodeValue();

                    XPathFactory xPathfactory = XPathFactory.newInstance();
                    XPath xpath = xPathfactory.newXPath();
                    String expression = "//artifactId[text()='" + artifactId + "']";

                    XPathExpression xPathExpr = xpath.compile(expression);

                    Node dependencyAlreadyExist = (Node) xPathExpr.evaluate(pomDocument, XPathConstants.NODE);

                    if (dependencyAlreadyExist != null) {
                        Node dependencies = pomDocument.getElementsByTagName("dependencies").item(0);
                        dependencies.removeChild(dependencyAlreadyExist.getParentNode());
                    }

                    Node importedNode = pomDocument.importNode(addTheseDependencies.item(n), true);
                    dependenciesNode.appendChild(importedNode);

                }
            }
            removeEmptyLinesfromDOM(pomDocument);
            writeToPOM(pomDocument);

        } catch (ParserConfigurationException | SAXException | XPathExpressionException | IOException ex) {
            Exceptions.printStackTrace(ex);
        }
    }

    /**
     *
     * @param addTheseRepositories
     */
    public void addMavenRepositoriesInProject(NodeList addTheseRepositories) {
        String mavenPomXmlfile = projectDir.getPath() + File.separator + AsposeConstants.MAVEN_POM_XML;

        try {
            Document pomDocument = getXmlDocument(mavenPomXmlfile);

            Node repositoriesNode = pomDocument.getElementsByTagName("repositories").item(0);

            if (addTheseRepositories != null && addTheseRepositories.getLength() > 0) {
                for (int n = 0; n < addTheseRepositories.getLength(); n++) {
                    String repositoryId = addTheseRepositories.item(n).getFirstChild().getNextSibling().getFirstChild().getNodeValue();

                    XPathFactory xPathfactory = XPathFactory.newInstance();
                    XPath xpath = xPathfactory.newXPath();
                    String expression = "//id[text()='" + repositoryId + "']";

                    XPathExpression xPathExpr = xpath.compile(expression);

                    Boolean repositoryAlreadyExist = (Boolean) xPathExpr.evaluate(pomDocument, XPathConstants.BOOLEAN);

                    if (!repositoryAlreadyExist) {
                        Node importedNode = pomDocument.importNode(addTheseRepositories.item(n), true);
                        repositoriesNode.appendChild(importedNode);
                    }

                }
            }
            removeEmptyLinesfromDOM(pomDocument);
            writeToPOM(pomDocument);

        } catch (XPathExpressionException | SAXException | ParserConfigurationException | IOException ex) {
            Exceptions.printStackTrace(ex);
        }
    }

    /**
     *
     * @param pomDocument
     * @throws IOException
     */
    public void writeToPOM(Document pomDocument) throws IOException {

        FileObject projectRoot = FileUtil.toFileObject(projectDir);
        FileObject fo = FileUtil.createData(projectRoot, AsposeConstants.MAVEN_POM_XML);
        try (OutputStream out = fo.getOutputStream()) {
            XMLUtil.write(pomDocument, out, "UTF-8");
        }
    }

    /**
     *
     * @param mavenPomXmlfile
     * @param excludeGroup
     * @return
     */
    public NodeList getDependenciesFromPOM(String mavenPomXmlfile, String excludeGroup) {

        try {

            Document pomDocument = getXmlDocument(mavenPomXmlfile);

            XPathFactory xPathfactory = XPathFactory.newInstance();
            XPath xpath = xPathfactory.newXPath();
            String expression = "//dependency[child::groupId[text()!='" + excludeGroup + "']]";
            XPathExpression xPathExpr = xpath.compile(expression);
            NodeList nl = (NodeList) xPathExpr.evaluate(pomDocument, XPathConstants.NODESET);
            if (nl != null && nl.getLength() > 0) {
                return nl;
            }
        } catch (IOException | ParserConfigurationException | SAXException | XPathExpressionException e) {
            Exceptions.printStackTrace(e);
        }
        return null;
    }

    /**
     *
     * @param mavenPomXmlfile
     * @param excludeURL
     * @return
     */
    public NodeList getRepositoriesFromPOM(String mavenPomXmlfile, String excludeURL) {

        try {

            Document pomDocument = getXmlDocument(mavenPomXmlfile);

            XPathFactory xPathfactory = XPathFactory.newInstance();
            XPath xpath = xPathfactory.newXPath();
            String expression = "//repository[child::url[not(starts-with(.,'" + excludeURL + "'))]]";
            XPathExpression xPathExpr = xpath.compile(expression);
            NodeList nl = (NodeList) xPathExpr.evaluate(pomDocument, XPathConstants.NODESET);
            if (nl != null && nl.getLength() > 0) {
                return nl;
            }
        } catch (IOException | ParserConfigurationException | SAXException | XPathExpressionException e) {
            Exceptions.printStackTrace(e);
        }
        return null;
    }

    private void removeEmptyLinesfromDOM(Document doc) throws XPathExpressionException {
        XPath xp = XPathFactory.newInstance().newXPath();
        NodeList nl = (NodeList) xp.evaluate("//text()[normalize-space(.)='']", doc, XPathConstants.NODESET);

        for (int i = 0; i < nl.getLength(); ++i) {
            Node node = nl.item(i);
            node.getParentNode().removeChild(node);
        }
    }

    /**
     *
     * @param asposeComponent
     * @param top
     * @param panel
     */
    public void populateExamplesTree(AsposeJavaAPI asposeComponent, CustomMutableTreeNode top, AsposeExamplePanel panel) {
        String examplesFullPath = asposeComponent.getLocalRepositoryPath() + File.separator + AsposeConstants.GITHUB_EXAMPLES_SOURCE_LOCATION;
        File directory = new File(examplesFullPath);
        panel.getExamplesTree().removeAll();
        top.setExPath(examplesFullPath);
        Queue<Object[]> queue = new LinkedList<>();
        queue.add(new Object[]{null, directory});

        while (!queue.isEmpty()) {
            Object[] _entry = queue.remove();
            File childFile = ((File) _entry[1]);
            CustomMutableTreeNode parentItem = ((CustomMutableTreeNode) _entry[0]);
            if (childFile.isDirectory()) {
                if (parentItem != null) {
                    CustomMutableTreeNode child = new CustomMutableTreeNode(FormatExamples.formatTitle(childFile.getName()));
                    child.setExPath(childFile.getAbsolutePath());
                    child.setFolder(true);
                    parentItem.add(child);
                    parentItem = child;
                } else {
                    parentItem = top;
                }
                for (File f : childFile.listFiles()) {
                    queue.add(new Object[]{parentItem, f});
                }
            } else if (childFile.isFile()) {

                String title = FormatExamples.formatTitle(childFile.getName());
                CustomMutableTreeNode child = new CustomMutableTreeNode(title);
                child.setFolder(false);
                parentItem.add(child);

            }
        }

    }
}
