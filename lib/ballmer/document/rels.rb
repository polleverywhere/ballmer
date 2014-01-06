module Ballmer
  class Document
    # CRUD and resolve relative documents to a part. These map to .xml.rel documents
    # in the MS Office document format.
    class Rels < Part
      attr_reader :path, :doc

      # TODO - Refactor the part_path business out here.
      def initialize(doc, path, part_path)
        super doc, path
        @part_path = Pathname.new(part_path)
      end

      # Return a list of target paths given a type.
      def targets(type)
        xml.xpath("//xmlns:Relationship[@Type='#{type}']").map{ |n| Pathname.new(n['Target']) }
      end

      # TODO 
      # Returns the rID of the part.
      def id(part)
        rel(part)['Id']
      end

      # Append a part to a rel so that we can extract an ID from it, and be
      # really cool like that.
      def append(part)
        return nil if exists? part

        xml.at_xpath('/xmlns:Relationships').tap do |relationships|
          relationships << Nokogiri::XML::Node.new("Relationship", xml).tap do |n|
            n['Id'] = next_id
            n['Type'] = part.class::REL_TYPE
            # Rels require a strange path... still haven't quite figured it out but I need to.
            n['Target'] = rel_path(part)
          end
        end
        commit
      end

      # TODO
      # Check if the part exists
      def exists?(part)
        !! rel(part)
      end

      # Create a Rels class from a given part.
      def self.from(part)
        new part.doc, rels_path(part.path), part.path
      end

      # Resolve the default rels asset for a given part path.
      def self.rels_path(part_path)
        Pathname.new(part_path).join('../_rels', Pathname.new(part_path).sub_ext('.xml.rels').basename)
      end
  
      private

      def rel(part)
        xml.at_xpath("/xmlns:Relationships/xmlns:Relationship[@Type='#{part.class::REL_TYPE}' and @Target='#{rel_path(part)}']")
      end

      # TODO - This feels dirty, dropping into kinda sorta paths (instead of parts). Refactor
      # this so that we're only dealing with parts up in here. Use Part#relative_path_from.
      def rel_path(rel_part)
        # I think the rel_part.path bidness is not returning and absolute path. Fix and maybe this will work (and
        # the weird + '..' won't be needed).
        rel_part.path.relative_path_from(@part_path + '..')
      end

      # TODO - Figure out how to make this more MS idiomatic up 9->10 instead of incrementing
      # the character....
      def next_id
        xml.xpath('//xmlns:Relationship[@Id]').map{ |n| n['Id'] }.sort.last.succ
      end
    end
  end
end
