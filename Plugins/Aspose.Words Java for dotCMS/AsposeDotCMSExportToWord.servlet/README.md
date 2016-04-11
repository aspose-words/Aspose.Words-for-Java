
README
------

Export to Word Plugin for dotCMS allow users to export online content into Word Processing document using Aspose.Words for Java. It dynamically exports the content of the Web page to a Word Processing document and then automatically downloads the file to the disk location selected by the user in just couple of seconds. The generated Word processing document can then be opened using any Word Processing Application such as Microsoft Word or Apache OpenOffice etc.

How to build this example
-------------------------

To install all you need to do is build the JAR. to do this run
./gradlew jar
This will build a jar in the build/libs directory

1. To install this bundle:

Copy the bundle jar file inside the Felix OSGI container (dotCMS/felix/load).
OR
Upload the bundle jar file using the dotCMS UI (CMS Admin->Dynamic Plugins->Upload Plugin).

Please add the following 2 packages by changing the file: dotCMS/WEB-INF/felix/osgi-extra.conf or using the dotCMS UI (System -> Dynamic Plugins -> Exported Packages).

i. javax.xml.stream
ii. javax.xml.namespace

2. To uninstall this bundle:

Remove the bundle jar file from the Felix OSGI container (dotCMS/felix/load).
OR
Undeploy the bundle using the dotCMS UI (CMS Admin->Dynamic Plugins->Undeploy).