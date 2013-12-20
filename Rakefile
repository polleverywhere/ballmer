# Uncomment if we want to make this public. Until then, we will
# distribute this as a gem from Github private repo.
#
# require "bundler/gem_tasks"

namespace :fixtures do
  desc "Unpack all of the fixtures for inspection with a diff tool"
  task :unpack do
    Dir['./fixtures/*.pptx'].each{ |f| sh "bin/unpack #{f}" }
  end
end