package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import java.util.UUID;

public class SignPersonTestClass {
    private UUID mPersonId;
    private String mName;
    private String mPosition;
    private byte[] mImage;

    public SignPersonTestClass() {
    }

    public SignPersonTestClass(final UUID guid, final String name, final String position, final byte[] image) {
        setPersonId(guid);
        setName(name);
        setPosition(position);
        setImage(image);
    }

    public void setPersonId(final UUID value) {
        mPersonId = value;
    }

    public void setName(final String value) {
        mName = value;
    }

    public void setPosition(final String value) {
        mPosition = value;
    }

    public void setImage(final byte[] value) {
        mImage = value;
    }

    public UUID getPersonId() {
        return mPersonId;
    }

    public String getName() {
        return mName;
    }

    public String getPosition() {
        return mPosition;
    }

    public byte[] getImage() {
        return mImage;
    }
}
