# A sample Guardfile
# More info at https://github.com/guard/guard#readme

guard :rspec do
  watch(%r{^spec/.+_spec\.rb$})
  watch(%r{^lib/(.+)\.rb$})     { |m| "spec/lib/#{m[1]}_spec.rb" }
  watch('spec/spec_helper.rb')  { "spec" }

  # Run the base spec since that's where a bulk of the tests hang out.
  watch(%r{^lib/ballmer/(.+?)/.+\.rb$})   { |m| "spec/lib/ballmer/#{m[1]}_spec.rb" }
end
