require "power_point/version"
require "zip"
require "nokogiri"

module PowerPoint
  # Your code goes here...
  class Presentation
    CONTENT_TYPES_PATH = "[Content_Types].xml"

    attr_reader :zip_file

    def initialize(zip_file)
      @zip_file = zip_file
    end

    # Return an array of slides.
    def slides
      slide_parts
    end

    def self.open(path)
      new Zip::File.open(path, Zip::File::CREATE)
    end

  private
    # Read and parse the contents of a zip entry.
    def xml(file)
      Nokogiri::XML zip_file.read file
    end

    # Reads from the [Content_Types].xml file the paths for the slide
    #   <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    def slide_parts
      parts(Slide::CONTENT_TYPE)
    end

    # Reads from the [Content_Types].xml file to grab paths fro the slide
    #   <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
    def slide_note_parts
      parts(SlideNote::CONTENT_TYPE)
    end

    # Query the [Content_Types].xml file to resolves paths to other parts in the zipfile.
    def parts(type)
      xml(CONTENT_TYPES_PATH).css("Override[ContentType='#{type}']").map{ |p| p['PartName'] }
    end
  end

  class Slide
    CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

    def initialize(slide, notes)
      @slide, @notes = slide, notes
    end

    def read(entry)
    end
  end

  class SlideNote
    CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
  end
end
