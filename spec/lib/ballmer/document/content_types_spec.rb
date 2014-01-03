require 'spec_helper'

describe Ballmer::Document::ContentTypes do
  subject     { read_presentation('presentation3.pptx') }
  let(:tags)  { Ballmer::Presentation::Tags.new(subject, 'tags/tag1.xml') }

  describe "#exists?" do
    it "should detect if part is in XML" do
      subject.content_types.should exist(subject.slides.first)
    end

    it "should detect if part is not in XML" do
      subject.content_types.should_not exist(tags)
    end
  end

  describe "#append" do
    before(:each) do
      subject.content_types.append tags
    end

    let(:nodes) do
      subject.content_types.xml.xpath("//xmlns:Override[@ContentType='#{tags.class::CONTENT_TYPE}' and @PartName='#{tags.path}']")
    end

    it "should write part to XML" do
      nodes.should have(1).item
    end

    it "should not write duplicate parts to XML" do
      subject.content_types.append tags
      nodes.should have(1).item
    end
  end
end
