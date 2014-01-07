require "bundler/gem_tasks"

namespace :fixtures do
  desc "Unpack all of the fixtures for inspection with a diff tool"
  task :unpack do
    Dir['./fixtures/*.pptx'].each{ |f| sh "bin/unpack #{f}" }
  end
end