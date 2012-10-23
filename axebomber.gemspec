# -*- encoding: utf-8 -*-

Gem::Specification.new do |s|
  s.name = %{axebomber}
  s.version  = "0.1.0"
  s.platform = Gem::Platform.new([nil, "java", nil])
  
  s.authors = ["kawasima"]
  s.description = %q{Excel parser for Hogan-style}
  s.email = ["kawasima1016@gmail.com"]
  s.files = [
  	*Dir["lib/**/*"].to_a
  ]
  ext = Dir.glob("ext/*ruby*")
  ext.delete(ext.detect{ |f| f =~ /jruby-complete/ })
  s.files += ext
  s.bindir = "bin"
  s.executables = []
  s.homepage = %q{https://github.com/kawasima/axebomber}
  s.rdoc_options = ["--main", "README.md"]
  s.require_paths = ['lib']
  s.summary = %q{maven support for ruby projects with gemspec, Gemfile}
end
