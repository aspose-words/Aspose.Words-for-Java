from src.quickstart.helloworld.python import HelloWorld
import jpype
import os.path

asposeapispath = os.path.join(os.path.abspath("./"), "lib")

print "You need to put your Aspose.Words for Java APIs .jars in this folder:\n"+asposeapispath

jpype.startJVM(jpype.getDefaultJVMPath(), "-Djava.ext.dirs=%s" % asposeapispath)

hw = HelloWorld()
hw.main()
