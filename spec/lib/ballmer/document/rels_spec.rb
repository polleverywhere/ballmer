require 'spec_helper'

describe Ballmer::Document::Rels do
  subject     { read_presentation('presentation3.pptx') }
  let(:tags)  { Ballmer::Presentation::Tags.new(subject, '/ppt/tags/tag1.xml') }
  let(:slide) { subject.slides.first }

  describe "#exists?" do
    it "should detect if part is in XML" do
      subject.presentation.rels.should exist(subject.slides.first)
    end

    it "should detect if part is not in XML" do
      subject.presentation.rels.should_not exist(tags)
    end
  end

  describe "#append" do
    before(:each) do
      slide.rels.append tags
    end

    let(:nodes) do
      slide.rels.xml.xpath("//xmlns:Relationship[@Type='#{tags.class::REL_TYPE}' and @Target='#{slide.relative_path_from(tags)}']")
    end

    it "should write part to XML" do
      nodes.should have(1).item
    end

    it "should not write duplicate parts to XML" do
      slide.rels.append tags
      nodes.should have(1).item
    end
  end

  describe "#id" do
    it "should get id" do
      # I just pulled this value out of the template.
      subject.presentation.rels.id(subject.slides.first).should == 'rId2'
    end
  end
end
