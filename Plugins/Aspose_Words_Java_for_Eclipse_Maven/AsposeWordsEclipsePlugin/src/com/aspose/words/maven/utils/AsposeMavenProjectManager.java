/*
 * The MIT License (MIT)
 *
 * Copyright (c) 1998-2016 Aspose Pty Ltd.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
package com.aspose.words.maven.utils;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.List;
import javax.swing.JOptionPane;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import com.aspose.words.maven.artifacts.Metadata;
import com.aspose.words.maven.artifacts.ObjectFactory;


public class AsposeMavenProjectManager {

	private File projectDir = null;

	private static final List<Metadata> asposeProjectMavenDependencies = new ArrayList<Metadata>();

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

	/**
	 *
	 * @return
	 */
	public File getProjectDir() {
		return projectDir;
	}

	public String getDependencyVersionFromPOM(URI projectDir, String dependencyName) {
		try {
			String mavenPomXmlfile = projectDir.getPath() + File.separator + AsposeConstants.MAVEN_POM_XML;

			if (new File(mavenPomXmlfile).exists()) {
				Document pomDocument = getXmlDocument(mavenPomXmlfile);

				XPathFactory xPathfactory = XPathFactory.newInstance();
				XPath xpath = xPathfactory.newXPath();
				String expression = "//version[ancestor::dependency/artifactId[text()='" + dependencyName + "']]";
				XPathExpression xPathExpr = xpath.compile(expression);
				NodeList nl = (NodeList) xPathExpr.evaluate(pomDocument, XPathConstants.NODESET);

				if (nl != null && nl.getLength() > 0) {
					return nl.item(0).getTextContent();
				}
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return null;
	}

	private Document getXmlDocument(String mavenPomXmlfile)
			throws ParserConfigurationException, SAXException, IOException {
		DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
		Document pomDocument = (Document) docBuilder.parse(mavenPomXmlfile);

		return pomDocument;
	}

	public String getAsposeHomePath() {
		return System.getProperty("user.home") + File.separator + "aspose" + File.separator;
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
			e.printStackTrace();
		}
		return null;
	}

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
					String artifactId = addTheseDependencies.item(n).getFirstChild().getNextSibling().getNextSibling()
							.getNextSibling().getFirstChild().getNodeValue();

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
			ex.printStackTrace();
		}
	}

	/**
	 *
	 * @return
	 */
	private boolean retrieveAsposeMavenDependencies() {
		try {
			getAsposeProjectMavenDependencies().clear();
			AsposeJavaAPI component = AsposeWordsJavaAPI.getInstance();
			Metadata productMavenDependency = getProductMavenDependency(component.get_mavenRepositoryURL());
			if (productMavenDependency != null) {
				getAsposeProjectMavenDependencies().add(productMavenDependency);
			}

		} catch (Exception rex) {
			rex.printStackTrace();
			return false;
		}
		return !getAsposeProjectMavenDependencies().isEmpty();
	}

	public void configureProjectMavenPOM(String groupId, String artifactId, String version) throws IOException {

		AsposeWordsJavaAPI.initialize(asposeMavenProjectManager);
		retrieveAsposeMavenDependencies();

		try {
			String mavenPomXmlfile = projectDir.getPath() + File.separator + AsposeConstants.MAVEN_POM_XML;
			Document doc = getXmlDocument(mavenPomXmlfile);

			Element root = doc.getDocumentElement();
			Node node = root.getElementsByTagName("groupId").item(0);
			node.setTextContent(groupId);

			node = root.getElementsByTagName("artifactId").item(0);
			node.setTextContent(artifactId);

			node = root.getElementsByTagName("version").item(0);
			node.setTextContent(version);

			updateProjectPom(doc);
			writeToPOM(doc);

		} catch (ParserConfigurationException | SAXException e) {
			e.printStackTrace();
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
		Metadata data = new Metadata();

		try {
			String productMavenInfo;
			productMavenInfo = readURLContents(productMavenRepositoryUrl + mavenMetaDataFileName);						
			 
			DocumentBuilder dBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
			Document doc = dBuilder.parse(new InputSource(new StringReader(productMavenInfo)));			
			XPath xPath = XPathFactory.newInstance().newXPath();															
			String groupId = XPathFactory.newInstance().newXPath().compile("//metadata/groupId").evaluate(doc);
			String artifactId = XPathFactory.newInstance().newXPath().compile("//metadata/artifactId").evaluate(doc);
			String version = XPathFactory.newInstance().newXPath().compile("//metadata/version").evaluate(doc);
			String latest = XPathFactory.newInstance().newXPath().compile("//metadata/versioning/latest").evaluate(doc);	 			
			
			data.setArtifactId(artifactId);
			data.setGroupId(groupId);
			data.setVersion(version);
			
			Metadata.Versioning ver = new Metadata.Versioning();
			ver.setLatest(latest);
			data.setVersioning(ver);
			
			String remoteArtifactFile = productMavenRepositoryUrl + data.getVersioning().getLatest() + "/"
					+ data.getArtifactId() + "-" + data.getVersioning().getLatest();
			

			if (!remoteFileExists(remoteArtifactFile + ".jar")) {
				AsposeConstants.println("Not Exists");
				data.setClassifier(getResolveSupportedJDK(remoteArtifactFile));
			} else {
				AsposeConstants.println("Exists");
			}
		} catch (Exception ex) {
			ex.printStackTrace();
			data = null;
		}
		return data;
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
			// HttpURLConnection.setInstanceFollowRedirects(false)
			HttpURLConnection con = (HttpURLConnection) new URL(URLName).openConnection();
			con.setRequestMethod("HEAD");
			return (con.getResponseCode() == HttpURLConnection.HTTP_OK);
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}

	/**
	 *
	 * @param ProductURL
	 * @return
	 */
	public String getResolveSupportedJDK(String ProductURL) {
		String supportedJDKs[] = { "jdk17", "jdk16", "jdk15", "jdk14", "jdk18" };
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
			e.printStackTrace();
		}
		return null;
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
					String repositoryId = addTheseRepositories.item(n).getFirstChild().getNextSibling().getFirstChild()
							.getNodeValue();

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
			ex.printStackTrace();
		}
	}

	/**
	 *
	 * @param pomDocument
	 * @throws IOException
	 */
	public void writeToPOM(Document pomDocument) throws IOException {
		try {
			TransformerFactory tFactory = TransformerFactory.newInstance();
			Transformer transformer = tFactory.newTransformer();
			DOMSource source = new DOMSource(pomDocument);

			StreamResult result = new StreamResult(
					new File(projectDir + File.separator + AsposeConstants.MAVEN_POM_XML));
			transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
			transformer.setOutputProperty(OutputKeys.METHOD, "xml");
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
			transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");

			transformer.transform(source, result);
		} catch (TransformerException e) {
			e.printStackTrace();
		}
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

	public static AsposeMavenProjectManager initialize(File prjDir) {
		asposeMavenProjectManager = new AsposeMavenProjectManager();
		asposeMavenProjectManager.projectDir = prjDir;
		return asposeMavenProjectManager;
	}

}