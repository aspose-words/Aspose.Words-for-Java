package DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects;

//ExStart:Manager
public class Manager
{
    private String mName;
    private int mAge;
    private byte[] mPhoto;
    private Iterable<Contract> mContracts;

    public String getName() { return mName; }
    public int getAge() { return mAge; }
    public byte[] getPhoto() { return mPhoto; }
    public Iterable<Contract> getContracts() { return mContracts; }

    public void setName(String value) { mName = value; }
    public void setAge(int value) { mAge = value; }
    public void setPhoto(byte[] value) { mPhoto = value; }
    public void setContracts(Iterable<Contract> value) { mContracts = value; }
}
//ExEnd:Manager
