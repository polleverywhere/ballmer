require "power_point/version"
require "nokogiri"

module PowerPoint
  autoload :Presentation, 'power_point/presentation'
  autoload :PPTX,         'power_point/pptx'

  # Deals with everything related to content paths.
  class ContentTypes
    PATH = "[Content_Types].xml"

    attr_reader :path, :pptx

    def initialize(pptx, path = PATH)
      @pptx, @path = pptx, path
    end

    # Get all of the parts for a given type
    # TODO - Have this return an enumerable of parts so we can fitler by part-type.
    def parts(type)
      xml.xpath("//xmlns:Override[@ContentType='#{type}']").map{ |n| n['PartName'] }
    end
    alias :[] :parts

    def xml
      pptx.xml(path)
    end
  end

  # Manages concerns around keeping slide and notesSlides files in
  # sync with an array of slides. These basically needs to trasnact
  # the slide\d+ and slideNote\d+ numbers to be in sync with an array.
  # Its a big, ugly ass complicated beast. Send your thank you cards to Bill Gates.
  class Slides
    include Enumerable

    def initialize(pptx)
      @pptx = pptx
    end

    def each(&block)
      # TODO - Do NOT read content-types, but read Rels instead (and move this type casting in there.)
      @pptx.content_types[Slide::CONTENT_TYPE].each { |path| block.call slide path }
    end

    def push(slide)
      n = to_a.size + 1
      # Paths within the zip file of new files we have to write.
      slide_path = Pathname.new("/ppt/slides/slide#{n}.xml")
      slide_rels_path = Pathname.new("/ppt/slides/_rels/slide#{n}.xml.rels")
      slide_notes_path = Pathname.new("/ppt/notesSlides/notesSlide#{n}.xml")
      slide_notes_rels_path = Pathname.new("/ppt/notesSlides/_rels/notesSlide#{n}.xml.rels")
      presentation_rels_path = Pathname.new("/ppt/_rels/presentation.xml.rels")
      presentation_path = Pathname.new("/ppt/presentation.xml")

      # Update ./ppt
      #   !!! CREATE !!!
      #   ./slides
      #     Create new files
      #       ./slide(\d+).xml file
      @pptx.copy slide_path, slide.path
      #       ./_rels/slide(\d+).xml.rels
      @pptx.copy slide_rels_path, slide.rels.path
      #   ./notesSlides
      #     Create new files
      #       ./notesSlide(\d+).xml file
      @pptx.copy slide_notes_path, slide.notes.path
      #       ./_rels/notesSlide(\d+).xml.rels
      @pptx.copy slide_notes_rels_path, slide.notes.rels.path
      
      #   !!! UPDATES !!!
      # Update the notes in the new slide to point at the new notes
      @pptx.edit_xml slide_rels_path do |xml|
        # TODO - Move this rel logic into the parts so that we don't have to repeat ourselves when calculating this stuff out.
        xml.at_xpath("//xmlns:Relationship[@Type='#{Notes::REL_TYPE}']")['Target'] = slide_notes_path.relative_path_from(slide_path.dirname)
      end

      # Update teh slideNotes reference to point at the new slide
      @pptx.edit_xml slide_notes_rels_path do |xml|
        xml.at_xpath("//xmlns:Relationship[@Type='#{Slide::REL_TYPE}']")['Target'] = slide_path.relative_path_from(slide_notes_path.dirname)
      end

      #   ./_rels/presentation.xml.rels
      #     Update Relationship ids
      #     Insert a new one slideRef
      @pptx.edit_xml presentation_rels_path do |xml|
        # Calucate the next id
        next_id = xml.xpath('//xmlns:Relationship[@Id]').map{ |n| n['Id'] }.sort.last.succ
        # TODO - Figure out how to make this more MS idiomatic up 9->10 instead of incrementing
        # the character....
        # Insert that into the slide and crakc open the presentation.xml file
        types = xml.at_xpath('/xmlns:Relationships')
        types << Nokogiri::XML::Node.new("Relationship", xml).tap do |n|
          n['Id'] = next_id
          n['Type'] = Slide::REL_TYPE
          n['Target'] = slide_path.relative_path_from(presentation_path.dirname)
        end
        #   ./presentation.xml
        #     Update attr
        #       p:notesMasterId
        #     Insert attr
        #       p:sldId, increment, etc.
        @pptx.edit_xml '/ppt/presentation.xml' do |xml|
          slides = xml.at_xpath('/p:presentation/p:sldIdLst')
          next_slide_id = slides.xpath('//p:sldId[@id]').map{ |n| n['id'] }.sort.last.succ
          slides << Nokogiri::XML::Node.new("p:sldId", xml).tap do |n|
            # TODO - Fix the ID that's jacked up.
            n['id'] = next_slide_id
            n['r:id'] = next_id
          end
        end
      end

      # Update ./[Content-Types].xml with new slide link and slideNotes link
      @pptx.edit_xml ContentTypes::PATH do |xml|
        types = xml.at_xpath('/xmlns:Types')
        types << Nokogiri::XML::Node.new("Override", xml).tap do |n|
          n['PartName'] = slide_path
          n['ContentType'] = Slide::CONTENT_TYPE
        end
        types << Nokogiri::XML::Node.new("Override", xml).tap do |n|
          n['PartName'] = slide_notes_path
          n['ContentType'] = Notes::CONTENT_TYPE
        end
      end

      # Great, that's all done, so lets return the slide eh?
      slide slide_path
    end

  private
    def slide(path)
      Slide.new(@pptx, path) 
    end
    # Reads from the [Content_Types].xml file the paths for the slide
    #   <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    def parts
      @pptx.content_types.parts Slide::CONTENT_TYPE
    end

    # Microsoft decided it would be cool to start at 1 instead of 0 
    # for the part indices, so this deals with that seperatly
    def next_number
      self.to_a.size + 1
    end
  end

  # Basic behavior of a part that we lift off of the [Content_Types].xml file.
  class Part
    attr_reader :path, :pptx

    def initialize(pptx, path)
      @pptx, @path = pptx, Pathname.new(path)
    end

    def xml
      @xml ||= pptx.xml(@path)
    end

    # Grab the rels file for this asset.
    def rels
      Rels.relative_to(self)
    end

    # Commit the part XML to the buffer.
    def commit
      @pptx.write path, xml.to_s
    end
  end

  # CRUD and resolve relative documents to a part. These map to .xml.rel documents
  # in the MS Office document format.
  class Rels
    attr_reader :path, :pptx

    def initialize(pptx, path, part_path)
      @pptx, @path, @part_path = pptx, path, part_path
    end

    # Return a list of target paths given a type.
    def targets(type)
      xml.xpath("//xmlns:Relationship[@Type='#{type}']").map{ |n| Pathname.new(n['Target']) }
    end

    def xml
      pptx.xml(@path)
    end

    # Create a Rels class from a given part.
    def self.relative_to(part)
      new part.pptx, rels_path(part.path), part.path
    end

    # Resolve the default rels asset for a given part path.
    def self.rels_path(part_path)
      Pathname.new(part_path).join('../_rels', Pathname.new(part_path).sub_ext('.xml.rels').basename)
    end
  end

  # Load a slide up in thar. 
  class Slide < Part
    # Key used to look up slides from [Content-Types].xml.
    CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".freeze

    # Key used to look up slides from .xml.rel documents
    REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide".freeze

    def notes
      # TODO - Move a type caster into rels based on content type like 
      # rels[Notes::REL_TYPE].first
      notes_path = rels.targets(Notes::REL_TYPE).first.expand_path(@path.dirname)
      Notes.new(@pptx, notes_path)
    end
  end

  class Notes < Part
    # TODO, there are three types of notes. We need to figure out 
    # how to resolve the slide number, notes, and whatever the hell else
    # the first note type is.

    # Key used to look up notes from [Content-Types].xml.
    CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml".freeze

    # Key used to look up notes from .xml.rel documents
    REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'.freeze

    # TODO - Generate/update the notes with a mark-down-ish heuristic, 
    # being that two newlines translate into the weird note formats of PPT slides.
    def body=(body)
      nodes_parser.parse(body)
      commit
    end

    def body
      nodes_parser.to_s
    end
    alias :to_s :body

  private
    def nodes_parser
      NotesParser.new(xml)
    end
  end

  class NotesParser
    # Make this read the language set by the OS or from a configuration.
    DEFAULT_LANG = 'en-US'

    def initialize(xml)
      @xml = xml
    end

    # MSFT Thought it would be cool to drop a bunch of different bodies
    # and id attributes in here that don't link to anything, so lets go
    # loosey goosey on it and find the stupid "Notes Placeholder" content.
    def node
      @xml.at_xpath('
        //p:nvSpPr[
          p:cNvPr[starts-with(@name, "Notes Placeholder")]
        ]/following-sibling::p:txBody
      ')
    end

    def to_s
      node.xpath('.//a:t').map(&:text).join("\n\n")
    end

    # Parses a text file with newline breaks into "paragraphs" per whatever weird markup
    # the noteSlides is using. For now we're keeping this simple, no italics or other crazy stuff,
    # but this is the class that would be extended, changed, or swapped out in the future.
    def parse(body, lang = DEFAULT_LANG)
      body_pr = Nokogiri::XML::Node.new("p:txBody", @xml)
      # These should be blank... I don't know why, but they're always that way in the files.
      #   <a:bodyPr/>
      #   <a:lstStyle/>
      body_pr << Nokogiri::XML::Node.new("a:bodyPr", @xml)
      body_pr << Nokogiri::XML::Node.new("a:lstStyle", @xml)
      # TODO - Reject blank lines after we chomp 'em
      body.lines.map(&:chomp).reject{ |l| l == "" }.each do |line|
        #     <a:r>
        #       <a:rPr lang="en-US" dirty="0" smtClean="0"/>
        #       <a:t>Poll A</a:t>
        body_pr << Nokogiri::XML::Node.new("a:p", @xml).tap do |p|
          p << Nokogiri::XML::Node.new("a:r", @xml).tap do |r|
            r << Nokogiri::XML::Node.new("a:rPr", @xml).tap do |rpr|
              # PPT just wants this, k?
              rpr["lang"] = lang
              rpr["dirty"] = "0"
              rpr["smtClean"] = "0"
            end
            # This is where we finally inject content. w00.
            r << Nokogiri::XML::Node.new("a:t", @xml).tap do |t|
              t.content = line.chomp
            end
          end
        end
      end

      node.replace body_pr
    end
  end
end