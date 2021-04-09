package ApiExamples.TestData.TestBuilders;

// ********* THIS FILE IS AUTO PORTED *********

import ApiExamples.ApiExampleBase;
import java.awt.image.BufferedImage;
import com.aspose.ms.System.IO.Stream;
import javax.imageio.ImageIO;
import ApiExamples.TestData.TestClasses.ImageTestClass;


public class ImageTestBuilder extends ApiExampleBase
{
    private BufferedImage mImage;
    private Stream mImageStream;
    private byte[] mImageBytes;
    private String mImageString;

    public ImageTestBuilder()
    {
        mImage = ImageIO.read(getImageDir() + "Transparent background logo.png");            
        mImageStream = Stream.Null;
        mImageBytes = new byte[0];
        mImageString = "";
    }

    public ImageTestBuilder withImage(BufferedImage image)
    {
        mImage = image;
        return this;
    }

    public ImageTestBuilder withImageStream(Stream imageStream)
    {
        mImageStream = imageStream;
        return this;
    }

    public ImageTestBuilder withImageBytes(byte[] imageBytes)
    {
        mImageBytes = imageBytes;
        return this;
    }

    public ImageTestBuilder withImageString(String imageString)
    {
        mImageString = imageString;
        return this;
    }

    public ImageTestClass build()
    {
        return new ImageTestClass(mImage, mImageStream, mImageBytes, mImageString);
    }
}
