require 'spec_helper'

describe Ballmer::Presentation::Tags do
  let(:presentation) { read_presentation('presentation1.pptx') }

  context "tags" do
    subject { Ballmer::Presentation::Tags.build(presentation, 'ppt/tags/tags1.xml') }

    it "should set key" do
      subject['foo'] = 'bar'
      subject.xml.at_xpath("//p:tag[@name='foo']")['val'].should == 'bar'
    end

    it "should read key" do
      subject['foo'] = 'bar'
      subject['foo'].should == 'bar'
    end

    it "should not write duplicate keys" do
      subject['foo'] = 'bar'
      subject['foo'] = 'bar'
      subject.xml.xpath("//p:tag[@name='foo']").should have(1).item
    end
  end
end
