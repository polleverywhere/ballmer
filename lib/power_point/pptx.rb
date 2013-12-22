module PowerPoint
  # Deals with file concerns between higher-level classes like
  # Slides, Notes and file-system level work.
  class PPTX
    attr_reader :zip

    def initialize(zip)
      @zip = zip
      @original_files = (0...zip.num_files).map { |n| zip.get_name(n) }
    end

    def save
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

    # Open an XML document at the given path.
    def xml(path)
      Nokogiri::XML read path
    end

    # Modify XML within a block and write it back to the zip when done.
    def edit_xml(path, &block)
      write path, xml(path).tap(&block).to_s
    end

    # Write to the zip file at the given path.
    def write(path, buffer)
      entries[path(path)] = buffer
    end

    # Read the blog from the Zifile
    def read(path)
      entries[path(path)]
    end

    def entries
      # TODO - Move this out into a buffer and deal with "commits" at a part level.
      @entries ||= Hash.new do |h,k|
        k = path(k).to_s
        h[k] = if @original_files.include? k
          zip.fopen(k).read
        else
          ""
        end
      end
    end

    # Copy a file in the zip from a path to a path.
    def copy(target, source)
      write target, read(source)
    end

    def content_types
      ContentTypes.new(self)
    end

  private
    # Normalize the path and resolve relative paths, if given.
    def path(path)
      Pathname.new(path).expand_path('/').relative_path_from(Pathname.new('/'))
    end
  end
end