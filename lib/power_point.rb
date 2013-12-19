require "power_point/version"
require "zip"
require "nokogiri"

module PowerPoint
  # Your code goes here...
  class Presentation
    CONTENT_TYPES_PATH = "[Content_Types].xml"

    def initialize(zip_file)
      @file_handler = FileHandler.new(zip_file)
    end

    # Return an array of slides.
    def slides
      slide_parts.map { |path| Slide.new(@file_handler, path) }
    end

    # Open a PPTX file from the given path.
    def self.open(path)
      new Zip::File.open(path, Zip::File::CREATE)
    end

  private
    # Reads from the [Content_Types].xml file the paths for the slide
    #   <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    def slide_parts
      parts(Slide::CONTENT_TYPE)
    end
    
    # Query the [Content_Types].xml file to resolves paths to other parts in the zipfile.
    def parts(type)
      @file_handler.xml(CONTENT_TYPES_PATH).xpath("//xmlns:Override[@ContentType='#{type}']").map{ |p| p['PartName'] }
    end
  end

  # Deals with file concerns between the Presentation and all associated 
  # elements, like notes, slides, etc.
  class FileHandler
    attr_reader :zip_file

    def initialize(zip_file)
      @zip_file = zip_file
    end

    # Read the blog from the Zifile
    def read(path)
      zip_file.read path(path)
    end

    # Normalize the path and resolve relative paths, if given.
    def path(path)
      Pathname.new(path).expand_path('/').to_s.gsub(/^\//, '')
    end

    # Open an XML document at the given path.
    def xml(path)
      Nokogiri::XML read path
    end
  end

  # Load a slide up in thar.
  class Slide
    CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

    def initialize(presentation, path)
      @presentation, @path = presentation, Pathname.new(path)

      # Open the rels file for this slide to see what other assets are at play.
      rels_path = Pathname.new(path).join('../_rels', Pathname.new(path).sub_ext('.xml.rels').basename)
      rels = @presentation.xml rels_path
      notes_path = Pathname.new(rels.xpath("//xmlns:Relationship[@Type='#{Notes::TYPE}']").first['Target']).expand_path(@path.dirname)

      # Lets grab what we can, then start to resolve some more paths from the rels doc.
      @slide = @presentation.xml path
      @notes = @presentation.xml notes_path
    end

    def notes
      # TODO, there are three types of notes. We need to figure out 
      # how to resolve the slide number, notes, and whatever the hell else
      # the first note type is.

      # Grab each paragram, then each line within that pragraph.
      # @notes.xpath('//p:txBody/a:p').map{ |ap| ap.xpath('//a:t') }.join("\n\n")

      # For now, I'm just going to splat out the whole damn thing.
      @notes.xpath('//a:t').map(&:text).join("\n\n")
    end

  private
    # TODO - Extract the rel logic in the initializer into a class so that we can query for all/other
    # assets in the PPT doc.
    def rels
    end
  end

  class Notes
    TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
  end
end
