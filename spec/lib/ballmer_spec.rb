require 'spec_helper'

describe Ballmer do
  describe ".path" do
    it "should resolve path" do
      File.should exist(Ballmer.path('../spec/fixtures/presentation3.pptx'))
    end
  end
end