require "zip"

module PowerPoint
  # Represents a presentation that has many slides.
  class Presentation
    CONTENT_TYPES_PATH = "[Content_Types].xml"

    def initialize(zip)
      @pptx = PPTX.new(zip)
    end

    # Save the .pptx file to disk.
    def save
      @pptx.save
    end

    # Return an array of slides.
    def slides
      @slides ||= Slides.new(@pptx)
    end

    # Open a PPTX file from the given path.
    def self.open(path)
      new Zip::File.open(path)
    end
  end
end