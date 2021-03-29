// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.VbaProject;
import com.aspose.words.VbaModule;
import com.aspose.words.VbaModuleType;
import org.testng.Assert;
import com.aspose.words.VbaModuleCollection;
import com.aspose.words.VbaReferenceCollection;
import com.aspose.words.VbaReference;
import com.aspose.words.VbaReferenceType;
import com.aspose.ms.System.msString;


@Test
class ExVbaProject !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void createNewVbaProject() throws Exception
    {
        //ExStart
        //ExFor:VbaProject.#ctor
        //ExFor:VbaProject.Name
        //ExFor:VbaModule.#ctor
        //ExFor:VbaModule.Name
        //ExFor:VbaModule.Type
        //ExFor:VbaModule.SourceCode
        //ExFor:VbaModuleCollection.Add(VbaModule)
        //ExFor:VbaModuleType
        //ExSummary:Shows how to create a VBA project using macros.
        Document doc = new Document();

        // Create a new VBA project.
        VbaProject project = new VbaProject();
        project.setName("Aspose.Project");
        doc.setVbaProject(project);

        // Create a new module and specify a macro source code.
        VbaModule module = new VbaModule();
        module.setName("Aspose.Module");
        module.setType(VbaModuleType.PROCEDURAL_MODULE);
        module.setSourceCode("New source code");

        // Add the module to the VBA project.
        doc.getVbaProject().getModules().add(module);

        doc.save(getArtifactsDir() + "VbaProject.CreateVBAMacros.docm");
        //ExEnd

        project = new Document(getArtifactsDir() + "VbaProject.CreateVBAMacros.docm").getVbaProject();

        Assert.assertEquals("Aspose.Project", project.getName());

        VbaModuleCollection modules = doc.getVbaProject().getModules();

        Assert.assertEquals(2, modules.getCount());

        Assert.assertEquals("ThisDocument", modules.get(0).getName());
        Assert.assertEquals(VbaModuleType.DOCUMENT_MODULE, modules.get(0).getType());
        Assert.assertNull(modules.get(0).getSourceCode());

        Assert.assertEquals("Aspose.Module", modules.get(1).getName());
        Assert.assertEquals(VbaModuleType.PROCEDURAL_MODULE, modules.get(1).getType());
        Assert.assertEquals("New source code", modules.get(1).getSourceCode());
    }

    @Test
    public void cloneVbaProject() throws Exception
    {
        //ExStart
        //ExFor:VbaProject.Clone
        //ExFor:VbaModule.Clone
        //ExSummary:Shows how to deep clone a VBA project and module.
        Document doc = new Document(getMyDir() + "VBA project.docm");
        Document destDoc = new Document();

        VbaProject copyVbaProject = doc.getVbaProject().deepClone();
        destDoc.setVbaProject(copyVbaProject);

        // In the destination document, we already have a module named "Module1"
        // because we cloned it along with the project. We will need to remove the module.
        VbaModule oldVbaModule = destDoc.getVbaProject().getModules().get("Module1");
        VbaModule copyVbaModule = doc.getVbaProject().getModules().get("Module1").deepClone();
        destDoc.getVbaProject().getModules().remove(oldVbaModule);
        destDoc.getVbaProject().getModules().add(copyVbaModule);

        destDoc.save(getArtifactsDir() + "VbaProject.CloneVbaProject.docm");
        //ExEnd

        VbaProject originalVbaProject = new Document(getArtifactsDir() + "VbaProject.CloneVbaProject.docm").getVbaProject();

        Assert.assertEquals(copyVbaProject.getName(), originalVbaProject.getName());
        Assert.assertEquals(copyVbaProject.getCodePage(), originalVbaProject.getCodePage());
        Assert.assertEquals(copyVbaProject.isSigned(), originalVbaProject.isSigned());
        Assert.assertEquals(copyVbaProject.getModules().getCount(), originalVbaProject.getModules().getCount());

        for (int i = 0; i < originalVbaProject.getModules().getCount(); i++)
        {
            Assert.assertEquals(copyVbaProject.getModules().get(i).getName(), originalVbaProject.getModules().get(i).getName());
            Assert.assertEquals(copyVbaProject.getModules().get(i).getType(), originalVbaProject.getModules().get(i).getType());
            Assert.assertEquals(copyVbaProject.getModules().get(i).getSourceCode(), originalVbaProject.getModules().get(i).getSourceCode());
        }
    }

    //ExStart
    //ExFor:VbaReference
    //ExFor:VbaReference.LibId
    //ExFor:VbaReferenceCollection
    //ExFor:VbaReferenceCollection.Count
    //ExFor:VbaReferenceCollection.RemoveAt(int)
    //ExFor:VbaReferenceCollection.Remove(VbaReference)
    //ExFor:VbaReferenceType
    //ExSummary:Shows how to get/remove an element from the VBA reference collection.
    @Test
    public void removeVbaReference() throws Exception
    {
        final String BROKEN_PATH = "X:\\broken.dll";
        Document doc = new Document(getMyDir() + "VBA project.docm");
        
        VbaReferenceCollection references = doc.getVbaProject().getReferences();
        Assert.assertEquals(5 ,references.getCount());
        
        for (int i = references.getCount() - 1; i >= 0; i--)
        {
            VbaReference reference = doc.getVbaProject().getReferences().get(i);
            String path = getLibIdPath(reference);
            
            if (BROKEN_PATH.equals(path))
                references.removeAt(i);
        }
        Assert.assertEquals(4 ,references.getCount());
        
        references.remove(references.get(1));
        Assert.assertEquals(3 ,references.getCount());
 
        doc.save(getArtifactsDir() + "VbaProject.RemoveVbaReference.docm"); 
    }
 
    /// <summary>
    /// Returns string representing LibId path of a specified reference. 
    /// </summary>
    private static String getLibIdPath(VbaReference reference)
    {
        switch (reference.getType())
        {
            case VbaReferenceType.REGISTERED:
            case VbaReferenceType.ORIGINAL:
            case VbaReferenceType.CONTROL:
                return getLibIdReferencePath(reference.getLibId());
            case VbaReferenceType.PROJECT:
                return getLibIdProjectPath(reference.getLibId());
            default:
                throw new IllegalArgumentException();
        }
    }
 
    /// <summary>
    /// Returns path from a specified identifier of an Automation type library.
    /// </summary>
    private static String getLibIdReferencePath(String libIdReference)
    {
        if (libIdReference != null)
        {
            String[] refParts = msString.split(libIdReference, '#');
            if (refParts.length > 3)
                return refParts[3];
        }
 
        return "";
    }
 
    /// <summary>
    /// Returns path from a specified identifier of an Automation type library.
    /// </summary>
    private static String getLibIdProjectPath(String libIdProject)
    {
        return libIdProject != null ? libIdProject.substring(3) : "";
    }
    //ExEnd
}

