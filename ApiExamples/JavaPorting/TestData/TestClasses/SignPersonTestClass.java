package ApiExamples.TestData.TestClasses;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.System.Guid;


public class SignPersonTestClass
{
    public Guid getPersonId() { return mPersonId; }; public void setPersonId(Guid value) { mPersonId = value; };

    private Guid mPersonId;
    public String getName() { return mName; }; public void setName(String value) { mName = value; };

    private String mName;
    public String getPosition() { return mPosition; }; public void setPosition(String value) { mPosition = value; };

    private String mPosition;
    public byte[] getImage() { return mImage; }; public void setImage(byte[] value) { mImage = value; };

    private byte[] mImage;

    public SignPersonTestClass(Guid guid, String name, String position, byte[] image)
    {
        setPersonId(guid);
        setName(name);
        setPosition(position);
        setImage(image);
    }
}
