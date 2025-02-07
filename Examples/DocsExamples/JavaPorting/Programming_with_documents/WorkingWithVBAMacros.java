package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.VbaProject;
import com.aspose.words.VbaModule;
import com.aspose.words.VbaModuleType;
import com.aspose.ms.System.msConsole;
import com.aspose.words.VbaReferenceCollection;
import com.aspose.words.VbaReference;
import com.aspose.words.VbaReferenceType;
import com.aspose.ms.System.msString;


class WorkingWithVba extends DocsExamplesBase
{
    @Test
    public void createVbaProject() throws Exception
    {
        //ExStart:CreateVbaProject
        //GistId:d9bac4ed890f81ea3de392ecfeedbc55
        Document doc = new Document();

        VbaProject project = new VbaProject();
        project.setName("AsposeProject");
        doc.setVbaProject(project);

        // Create a new module and specify a macro source code.
        VbaModule module = new VbaModule();
        module.setName("AsposeModule");
        module.setType(VbaModuleType.PROCEDURAL_MODULE);
        module.setSourceCode("New source code");

        // Add module to the VBA project.
        doc.getVbaProject().getModules().add(module);

        doc.save(getArtifactsDir() + "WorkingWithVba.CreateVbaProject.docm");
        //ExEnd:CreateVbaProject
    }

    @Test
    public void readVbaMacros() throws Exception
    {
        //ExStart:ReadVbaMacros
        //GistId:d9bac4ed890f81ea3de392ecfeedbc55
        Document doc = new Document(getMyDir() + "VBA project.docm");

        if (doc.getVbaProject() != null)
        {
            for (VbaModule module : doc.getVbaProject().getModules())
            {
                System.out.println(module.getSourceCode());
            }
        }
        //ExEnd:ReadVbaMacros
    }

    @Test
    public void modifyVbaMacros() throws Exception
    {
        //ExStart:ModifyVbaMacros
        //GistId:d9bac4ed890f81ea3de392ecfeedbc55
        Document doc = new Document(getMyDir() + "VBA project.docm");

        VbaProject project = doc.getVbaProject();

        final String NEW_SOURCE_CODE = "Test change source code";
        project.getModules().get(0).setSourceCode(NEW_SOURCE_CODE);
        //ExEnd:ModifyVbaMacros
        
        doc.save(getArtifactsDir() + "WorkingWithVba.ModifyVbaMacros.docm");
        //ExEnd:ModifyVbaMacros
    }

    @Test
    public void cloneVbaProject() throws Exception
    {
        //ExStart:CloneVbaProject
        //GistId:d9bac4ed890f81ea3de392ecfeedbc55
        Document doc = new Document(getMyDir() + "VBA project.docm");
        Document destDoc = new Document(); { destDoc.setVbaProject(doc.getVbaProject().deepClone()); }

        destDoc.save(getArtifactsDir() + "WorkingWithVba.CloneVbaProject.docm");
        //ExEnd:CloneVbaProject
    }

    @Test
    public void cloneVbaModule() throws Exception
    {
        //ExStart:CloneVbaModule
        //GistId:d9bac4ed890f81ea3de392ecfeedbc55
        Document doc = new Document(getMyDir() + "VBA project.docm");
        Document destDoc = new Document(); { destDoc.setVbaProject(new VbaProject()); }
        
        VbaModule copyModule = doc.getVbaProject().getModules().get("Module1").deepClone();
        destDoc.getVbaProject().getModules().add(copyModule);

        destDoc.save(getArtifactsDir() + "WorkingWithVba.CloneVbaModule.docm");
        //ExEnd:CloneVbaModule
    }

    @Test
    public void removeVbaReferences() throws Exception
    {
        //ExStart:RemoveVbaReferences
        //GistId:d9bac4ed890f81ea3de392ecfeedbc55
        Document doc = new Document(getMyDir() + "VBA project.docm");

        // Find and remove the reference with some LibId path.
        final String BROKEN_PATH = "brokenPath.dll";
        VbaReferenceCollection references = doc.getVbaProject().getReferences();
        for (int i = references.getCount() - 1; i >= 0; i--)
        {
            VbaReference reference = doc.getVbaProject().getReferences().ElementAt(i);

            String path = getLibIdPath(reference);
            if (BROKEN_PATH.equals(path))
                references.removeAt(i);
        }

        doc.save(getArtifactsDir() + "WorkingWithVba.RemoveVbaReferences.docm");
        //ExEnd:RemoveVbaReferences
    }
    //ExStart:GetLibIdAndReferencePath
    //GistId:d9bac4ed890f81ea3de392ecfeedbc55
    /// <summary>
    /// Returns string representing LibId path of a specified reference. 
    /// </summary>
    private String getLibIdPath(VbaReference reference)
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
    /// <remarks>
    /// Please see details for the syntax at [MS-OVBA], 2.1.1.8 LibidReference. 
    /// </remarks>
    private String getLibIdReferencePath(String libIdReference)
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
    /// <remarks>
    /// Please see details for the syntax at [MS-OVBA], 2.1.1.12 ProjectReference. 
    /// </remarks>
    private String getLibIdProjectPath(String libIdProject)
    {
        return (libIdProject != null) ? libIdProject.substring(3) : "";
    }
    //ExEnd:GetLibIdAndReferencePath
}
