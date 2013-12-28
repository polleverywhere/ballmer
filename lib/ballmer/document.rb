module Ballmer
  # Deals with file concerns between higher-level classes like
  # Slides, Notes and file-system level work.
  class Document
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
      Document::ContentTypes.new(self)
    end

    private
    # Normalize the path and resolve relative paths, if given.
    def path(path)
      Pathname.new(path).expand_path('/').relative_path_from(Pathname.new('/'))
    end
  end

  class Document
    # Deals with everything related to content paths.
    class ContentTypes
      PATH = "[Content_Types].xml"

      attr_reader :path, :doc

      def initialize(doc, path = PATH)
        @doc, @path = doc, path
      end

      # Get all of the parts for a given type
      # TODO - Have this return an enumerable of parts so we can fitler by part-type.
      def parts(type)
        xml.xpath("//xmlns:Override[@ContentType='#{type}']").map{ |n| n['PartName'] }
      end
      alias :[] :parts

      def xml
        doc.xml(path)
      end
    end

    # Basic behavior of a part that we lift off of the [Content_Types].xml file.
    class Part
      attr_reader :path, :doc

      def initialize(doc, path)
        @doc, @path = doc, Pathname.new(path)
      end

      def xml
        @xml ||= doc.xml(@path)
      end

      # Grab the rels file for this asset.
      def rels
        Rels.relative_to(self)
      end

      # Commit the part XML to the buffer.
      def commit
        @doc.write path, xml.to_s
      end
    end

    # CRUD and resolve relative documents to a part. These map to .xml.rel documents
    # in the MS Office document format.
    class Rels
      attr_reader :path, :doc

      def initialize(doc, path, part_path)
        @doc, @path, @part_path = doc, path, part_path
      end

      # Return a list of target paths given a type.
      def targets(type)
        xml.xpath("//xmlns:Relationship[@Type='#{type}']").map{ |n| Pathname.new(n['Target']) }
      end

      def xml
        doc.xml(@path)
      end

      # Create a Rels class from a given part.
      def self.relative_to(part)
        new part.doc, rels_path(part.path), part.path
      end

      # Resolve the default rels asset for a given part path.
      def self.rels_path(part_path)
        Pathname.new(part_path).join('../_rels', Pathname.new(part_path).sub_ext('.xml.rels').basename)
      end
    end
  end
end