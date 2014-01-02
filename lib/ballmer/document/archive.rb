require "zipruby"

module Ballmer
  class Document
    # Abstraction that sits on top of ZipRuby because the original
    # lib API is a bit cumbersom to use directly.
    class Archive
      attr_reader :zip

      def initialize(zip)
        @zip = zip
        @original_files = (0...zip.num_files).map { |n| zip.get_name(n) }
      end

      # Save the office XML file to the file.
      def commit
        # TODO
        # Update ./docProps
        #   app.xml slides, notes, counts, etc
        #   core.xml Times
        entries.each do |path, buffer|
          path = path.to_s
          if @original_files.include? path
            @zip.replace_buffer path, buffer
          else
            @zip.add_buffer path, buffer
          end
        end
        @zip.commit
      end

      # Write to the zip file at the given path.
      def write(path, buffer)
        entries[self.class.path(path)] = buffer
      end

      # Read the blog from the Zifile
      def read(path)
        entries[self.class.path(path)]
      end

      # Copy a file in the zip from a path to a path.
      def copy(target, source)
        write target, read(source)
      end

      # Enumerates all of the entries in the zip file. Key
      # is the path of the file, and the value is the contents.
      def entries
        @entries ||= Hash.new do |h,k|
          k = self.class.path(k).to_s
          h[k] = if @original_files.include? k
            zip.fopen(k).read
          else
            ""
          end
        end
      end

      private
      # Normalize the path and resolve relative paths, if given.
      def self.path(path)
        Pathname.new(path).expand_path('/').relative_path_from(Pathname.new('/'))
      end

      # Open an XML office file from the given path.
      def self.open(path)
        new Zip::Archive.open(path, Zip::TRUNC)
      end

      # Read zip data from a bufffer. Very useful when you want to load a template 
      # into a server environment, modify, and serve up without writing to disk.
      def self.read(data)
        new Zip::Archive.open_buffer(data)
      end
    end
  end
end