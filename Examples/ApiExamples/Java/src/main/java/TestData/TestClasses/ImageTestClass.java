package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import java.awt.image.BufferedImage;
import java.io.FileInputStream;

public class ImageTestClass {
    private BufferedImage mImage;
    private FileInputStream mImageStream;
    private byte[] mImageBytes;
    private String mImageString;

    public ImageTestClass() {
    }

    public ImageTestClass(final BufferedImage image, final FileInputStream imageStream, final byte[] imageBytes, final String imageString) {
        setImage(image);
        setImageStream(imageStream);
        setImageBytes(imageBytes);
        setImageString(imageString);
    }

    public void setImage(final BufferedImage value) {
        mImage = value;
    }

    public void setImageStream(final FileInputStream value) {
        mImageStream = value;
    }

    public void setImageBytes(final byte[] value) {
        mImageBytes = value;
    }

    public void setImageString(final String value) {
        mImageString = value;
    }

    public BufferedImage getImage() {
        return mImage;
    }

    public FileInputStream getImageStream() {
        return mImageStream;
    }

    public byte[] getImageBytes() {
        return mImageBytes;
    }

    public String getImageString() {
        return mImageString;
    }
}
