require "zipruby"

module Ballmer
  # Represents a presentation that has many slides.
  class Presentation
    attr_reader :pptx

    def initialize(zip)
      @pptx = PPTX.new(zip)
    end

    # Save the .pptx file to disk.
    def save
      pptx.save
    end

    # Return an array of slides.
    def slides
      @slides ||= Slides.new(@pptx)
    end

    # Open a PPTX file from the given path.
    def self.open(path)
      new Zip::Archive.open(path, Zip::TRUNC)
    end
  end
end