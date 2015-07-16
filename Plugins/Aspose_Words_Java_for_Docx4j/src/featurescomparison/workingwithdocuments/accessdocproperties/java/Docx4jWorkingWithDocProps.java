/*
 *  Copyright 2007-2008, Plutext Pty Ltd.
 *   
 *  This file is part of docx4j.

    docx4j is licensed under the Apache License, Version 2.0 (the "License"); 
    you may not use this file except in compliance with the License. 

    You may obtain a copy of the License at 

        http://www.apache.org/licenses/LICENSE-2.0 

    Unless required by applicable law or agreed to in writing, software 
    distributed under the License is distributed on an "AS IS" BASIS, 
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
    See the License for the specific language governing permissions and 
    limitations under the License.

	NOTICE: ORIGINAL FILE MODIFIED
*/

package featurescomparison.workingwithdocuments.accessdocproperties.java;

import java.util.List;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.samples.AbstractSample;

public class Docx4jWorkingWithDocProps extends AbstractSample 
{
	public static void main(String[] args) throws Exception 
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/accessdocproperties/data/";
		
		inputfilepath = dataPath + "document.docx";

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputfilepath));

		// Let's look at the core properties
		org.docx4j.openpackaging.parts.DocPropsCorePart docPropsCorePart = wordMLPackage.getDocPropsCorePart();
		org.docx4j.docProps.core.CoreProperties coreProps = (org.docx4j.docProps.core.CoreProperties)docPropsCorePart.getJaxbElement();

		// Title of the document
		// Note: Word for Mac 2010 doesn't set title
		String title = "Missing";
		List<String> list = coreProps.getTitle().getValue().getContent();
		if (list.size() > 0) 
		{
			title = list.get(0);
		}
		System.out.println("'dc:title' is " + title);

		// Extended properties
		org.docx4j.openpackaging.parts.DocPropsExtendedPart docPropsExtendedPart = wordMLPackage.getDocPropsExtendedPart();
		org.docx4j.docProps.extended.Properties extendedProps = (org.docx4j.docProps.extended.Properties)docPropsExtendedPart.getJaxbElement(); 

		// Document creator Application
		System.out.println("'Application' is " + extendedProps.getApplication() + " v." + extendedProps.getAppVersion());

		// Custom properties
		org.docx4j.openpackaging.parts.DocPropsCustomPart docPropsCustomPart = wordMLPackage.getDocPropsCustomPart();
		if(docPropsCustomPart==null)
		{
			System.out.println("No Document Custom Properties.");
		} 
		else 
		{
			org.docx4j.docProps.custom.Properties customProps = (org.docx4j.docProps.custom.Properties)docPropsCustomPart.getJaxbElement();

			for (org.docx4j.docProps.custom.Properties.Property prop: customProps.getProperty() ) 
			{
				// At the moment, you need to know what sort of value it has.
				// Could create a generic Object getValue() method.
				if (prop.getLpwstr()!=null) 
				{
					System.out.println(prop.getName() + " = " + prop.getLpwstr());
				} 
				else 
				{
					System.out.println(prop.getName() + ": \n " + XmlUtils.marshaltoString(prop, true, Context.jcDocPropsCustom));
				}
			}
		}
		System.out.println("Done.");
	}
}