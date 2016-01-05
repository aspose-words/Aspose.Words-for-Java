from asposewords import Settings
from com.aspose.words import License

class ApplyLicense:

    def __init__(self):
        dataDir = Settings.dataDir + 'quickstart/'
        
        # This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
        # You can also use the additional overload to load a license from a stream, this is useful for instance when the
        # license is stored as an embedded resource
        try:
            license = License()
            license.setLicense("Aspose.Words.lic")
            print "License set successfully."
        except Exception as e:
            # We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license.
            print "There was an error setting the license." + e.getMessage()

if __name__ == '__main__':               
    ApplyLicense()