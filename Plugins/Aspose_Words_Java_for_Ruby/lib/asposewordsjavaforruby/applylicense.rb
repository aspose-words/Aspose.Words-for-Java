module Asposewordsjavaforruby
  module ApplyLicense
    def initialize()
        apply_license()
    end
    
    def apply_license()
        # This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
        # You can also use the additional overload to load a license from a stream, this is useful for instance when the
        # license is stored as an embedded resource
        license = Rjb::import('com.aspose.words.License').new()
        license.setLicense('Aspose.Words.lic')
    end

  end
end
