package ApiExamples.TestData.TestClasses;

// ********* THIS FILE IS AUTO PORTED *********

import java.awt.image.BufferedImage;
import com.aspose.ms.System.IO.Stream;


public class ImageTestClass
{
    public BufferedImage getImage() { return mImage; }; public void setImage(BufferedImage value) { mImage = value; };

    private BufferedImage mImage;        
    public Stream getImageStream() { return mImageStream; }; public void setImageStream(Stream value) { mImageStream = value; };

    private Stream mImageStream;
    public byte[] getImageBytes() { return mImageBytes; }; public void setImageBytes(byte[] value) { mImageBytes = value; };

    private byte[] mImageBytes;
    public String getImageString() { return mImageString; }; public void setImageString(String value) { mImageString = value; };

    private String mImageString;

    public ImageTestClass(BufferedImage image, Stream imageStream, byte[] imageBytes, String imageString)
    {
        setImage(image);
        setImageStream(imageStream);
        setImageBytes(imageBytes);
        setImageString(imageString);
    }
}
