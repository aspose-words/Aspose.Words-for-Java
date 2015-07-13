__author__ = 'fahadadeel'
import jpype
import os.path
from quickstart import AppendDocuments

asposeapispath = os.path.join(os.path.abspath("../../../"), "lib")

print "You need to put your Aspose.Words for Java APIs .jars in this folder:\n"+asposeapispath

jpype.startJVM(jpype.getDefaultJVMPath(), "-Djava.ext.dirs=%s" % asposeapispath)

ap = AppendDocuments();
ap.main('data/TestFile Out.docx','data/TestFile.Destination.doc','data/TestFile.Source.doc')
