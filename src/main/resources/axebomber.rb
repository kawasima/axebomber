if RUBY_PLATFORM =~ /java/
  require "java"
  Dir.chdir(File.dirname(__FILE__)) do
  	Dir.glob("*.jar").each do |jar|
  	  require jar
  	end
  end
elsif $VERBOSE
  warn "axebomber is only for use with JRuby"
end

module Axebomber
  VERSION = "0.1.0-SNAPSHOT"
  java_import "net.unit8.axebomber.manager.impl.ReadOnlyFileSystemBookManager"
end