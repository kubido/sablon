# -*- coding: utf-8 -*-
require 'test_helper'

class ImagesTest < Sablon::TestCase
  include Sablon::Test::Assertions

  def setup
    super
    @base_path = Pathname.new(File.expand_path('../', __FILE__))
    @fixtures_path = File.join(@base_path, 'fixtures')
    @output_path = File.join(@base_path, 'sandbox/images.docx')
  end

  def test_generate_document_with_images
    template = Sablon.template(File.join(@fixtures_path, 'images_template.docx'))
    images = [
      Sablon::Image.new('test1.jpg', File.open(File.join(@fixtures_path, 'test1.jpg'), 'rb') {|f| f.read}, nil),
      Sablon::Image.new('test2.jpg', File.open(File.join(@fixtures_path, 'test2.jpg'), 'rb') {|f| f.read}, nil),
      Sablon::Image.new('test3.jpg', File.open(File.join(@fixtures_path, 'test3.jpg'), 'rb') {|f| f.read}, nil)
    ]
    context = {
      :items => [
        { 'value' => 'Foo', 'image' => images[0] },
        { 'value' => 'Bar', 'image' => images[1] },
        { 'value' => 'Baz', 'image' => images[2] }
      ]
    }
    template.render_to_file @output_path, context, images

    assert_docx_equal File.join(@fixtures_path, 'images_sample.docx'), @output_path
  end
end
