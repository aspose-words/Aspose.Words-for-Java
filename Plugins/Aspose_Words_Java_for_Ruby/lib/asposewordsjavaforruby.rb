require_relative 'asposewordsjavaforruby/version'
require_relative 'asposewordsjavaforruby/asposewordsjava'
require 'logger'
require 'rjb'

module Asposewordsjavaforruby

  class << self
    attr_accessor :aspose_words_config
  end

  def initialize_aspose_words
    aspose_jars_dir = Asposewordsjavaforruby.aspose_words_config ? Asposewordsjavaforruby.aspose_words_config['jar_dir'] : nil  
    aspose_license_path = Asposewordsjavaforruby.aspose_words_config ? Asposewordsjavaforruby.aspose_words_config['license_path'] : nil
    jvm_args = Asposewordsjavaforruby.aspose_words_config ? Asposewordsjavaforruby.aspose_words_config['jvm_args'] : nil
    
    load_aspose_jars(aspose_jars_dir, jvm_args)
    load_aspose_license(aspose_license_path)
  end

  def load_aspose_license(aspose_license_path)
    if aspose_license_path && File.exist?(aspose_license_path)
      set_license(File.join(aspose_license_path))
    else
      logger = Logger.new(STDOUT)
      logger.level = Logger::WARN
      logger.warn('Using the non licensed aspose jar. Please specify path to your aspose license directory in config/aspose.yml file!')
    end
  end

  def load_aspose_jars(aspose_jars_dir, jvm_args)
    if aspose_jars_dir && File.exist?(aspose_jars_dir)
      jardir = File.join(aspose_jars_dir, '**', '*.jar')
    else
      jardir = File.join(File.dirname(File.dirname(__FILE__)), 'jars', '**', '*.jar')
    end

    if jvm_args
      args = jvm_args.split(' ') << '-Djava.awt.headless=true'
      logger = Logger.new(STDOUT)
      logger.level = Logger::DEBUG
      logger.debug("JVM args : #{args}")
      Rjb::load(classpath = Dir.glob(jardir).join(':'), jvmargs=args)
    else
      Rjb::load(classpath = Dir.glob(jardir).join(':'), jvmargs=['-Djava.awt.headless=true'])
    end

  end

  def input_file(file)
    Rjb::import('java.io.FileInputStream').new(file)
  end

  def set_license(aspose_license_file)
    begin
      fstream = input_file(aspose_license_file)
      license = Rjb::import('com.aspose.words.License').new()
      license.setLicense(fstream)
    rescue Exception => ex
      logger = Logger.new(STDOUT)
      logger.level = Logger::ERROR
      logger.error("Could not load the license file : #{ex}")
      fstream.close() if fstream
    end
  end

  def self.configure_aspose_words config
    Asposewordsjavaforruby.aspose_words_config = config
  end

end
