==========================================
Aspose.Words for Java Samples Read Me
==========================================

This package contains Sample Projects for Aspose.Words for Java.

Sample Projects are distributed separately from the Aspose.Words for Java download.

There are IntelliJ Idea project files (.ipr) provided for the samples which can be used to easily compile and run each sample.

How to Install the Samples
==========================================

Extract all of the samples to a folder called "Samples" at same location that the Aspose.Words binares were extracted to. The Samples folder should be extracted along side the "demos" and "lib" folder in order for the required references to be found. 

The lib folder contains product libraries, including Aspose.Words that are required to build the demos. Following the previous step allows referenced libraries to be located automatically when using the IntelliJ project files and the build scripts.


How to Run the Samples
==========================================

The Samples folder contains many different directories which each house a separate complete sample. The general structure of each sample contains:

- IntelliJ Project files (.ipr, .iml, .iws)
- Program.java (the source code for the sample).
- Data folder which contains the input documents and any source files. All generated output produced by the sample is also placed into this folder.
- Out folder (created when the sample is compiled and contains the compiled classes when the project is built).

You can choose one of several options to run the demos.

1)
Open the IntelliJ project file using the IntelliJ IDEA Java IDE of the sample you wish to run. Click on "Run" menu and choose one of the following menu items:

- Run
- Debug

2)

The samples are shipped with scripts found at the root of the Samples folder: "Compile All.bat" and "Run All.bat". These allow all of the sample projects within the folder to be built and run automatically under Windows. 

The Java compiler and runtime are found by using environment variables. Therefore the JAVA_HOME variable should be correctly defined and the appropriate directory pointing to a JDK and java.exe present in the PATH variable in orer for the scripts to work.

The compile script will attempt to compile all of the samples. A message within the console displays which sample is being compiled and if it was successfully built or not. If an error occurred then the full stack trace is appended to a log file called "BuildLog.txt" in the current directory.

The run script will attempt to run all of the samples sequentially. The output of the sample is displayed on the console if there is any. Most output of samples do not display anything on the console as they only generate documents which can be found in the Data folder of each sample. If there are any run time errors then the full stack trace is displayed to the console.

3)
Create a new project in your favourite Java IDE and choose to create a new project from existing sources. 

- Include all files within the folder of the sample that is being compiled. 
- A library reference to Aspose.Words.jar which is found within the lib folder must be added to the project. 
- Depending on the sample being compiled, a reference to other Aspose.Words libraries may also be required. These are found under the demos-only-lib folder inside the lib folder.
- The entry point for each sample is always found at "Program.Main".


Software Requirements
==========================================

- Aspose.Words for Java 10.6.0 or later
- An additional library: Aspose.Network for Java 2.0.1 or later. This is used to demonstrate integration of Aspose.Words with other Aspose libraries.
- The TestNG plugin must be enabled in IntelliJ IDEA in order to properly run the test code found in Examples.
- For samples which use a graphical interface the "UI Designer" plugin must be installed and enabled. This can be found under File -> Settings -> Plugins. "UI Designer" needs to be present and enabled on the installed plugins list.

- Example code which demonstrates conversion of a document to image requires the appropriate ImageIO codecs installed on the system. If one of these codecs is missing it will present itself as a runtime error such as: "java.lang.IllegalStateException: Cannot find an ImageIO writer for the specified format: bmp". See below for details on how to download and install ImageIO.

- Java comes with most standard ImageIO codecs during most installations, however in some situations missing codecs must be downloaded manually:
  - The BMP codec is missing from Java 1.4.
  - The TIFF codec is missing from Java 1.5.

Information on how to install ImageIO codecs and the binaries can be found at the following site: http://download.java.net/media/jai-imageio/builds/release/1.1/


Running the Samples under Linux and MacOS
==========================================

The samples should be compiled and run using either of the Sun-JDK or Open-JDK. There should be no extra requirements as long as the standard Java libraries are present.

There is a known limitation when running the samples in Linux. Some samples use a Microsoft Access database to read data to merge into the documents. Currently there is no known "free" driver for reading MDB databases in Linux that will work up to the standards required by the samples.Therefore this functionality is missing a some of the samples cannnot be executed.

We are looking into resolving this issue by changing the data source used in the these particular samples.


http://www.aspose.com/
Copyright (c) 2001-2011 Aspose Pty Ltd. All Rights Reserved.