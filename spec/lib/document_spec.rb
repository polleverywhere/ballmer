require 'spec_helper'

describe Ballmer::Document do
  subject { Ballmer::Document.read File.read "./fixtures/Presentation3.pptx" }

  context "rels" do
    it "should resolve"
  end

  context "content_types"
  context "zip"
end