# -*- coding: utf-8 -*-
module Sablon
  class Processor
    RELATIONSHIPS_NS_URI = 'http://schemas.openxmlformats.org/package/2006/relationships'
    PICTURE_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    MAIN_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

    def self.process(xml_node, context, properties = {})
      processor = new(parser)
      stringified_context = Hash[context.map {|k, v| [k.to_s, v] }]
      processor.manipulate xml_node, stringified_context
      processor.write_properties xml_node, properties if properties.any?
      xml_node
    end

    def self.process_rels(xml_node, images)
      next_id = next_rel_id(xml_node)
      relationships = xml_node.at_xpath('r:Relationships', 'r' => RELATIONSHIPS_NS_URI)
      images.each do |image|
        relationships.add_child("<Relationship Id='rId#{next_id}' Type='#{IMAGE_TYPE}' Target='media/#{image.name}'/>")
        image.rid = next_id
        next_id += 1
      end
      xml_node
    end

    def self.parser
      @parser ||= Sablon::Parser::MailMerge.new
    end

    def self.next_rel_id(xml_node)
      max = 0
      xml_node.xpath('r:Relationships/r:Relationship', 'r' => RELATIONSHIPS_NS_URI).each do |n|
        id = n.attributes['Id'].to_s[3..-1].to_i
        max = id if id > max
      end
      max + 1
    end

    def self.remove_final_blank_page(xml_node)
      children = xml_node.xpath('/w:document/w:body/*')
      found_last = false
      children.reverse.each do |child|
        if found_last
          if child.name == 'p' && child.namespace.prefix == 'w'
            page_break = child.xpath("w:r/w:br[@w:type='page']")
            page_break.remove unless page_break.nil?
            break
          end
        elsif child.name == 'sectPr' && child.namespace.prefix == 'w'
          found_last = true
        end
      end
      xml_node
    end

    def initialize(parser)
      @parser = parser
    end

    def manipulate(xml_node, context)
      operations = build_operations(@parser.parse_fields(xml_node))
      operations.each do |step|
        step.evaluate context
      end
      xml_node
    end

    def write_properties(xml_node, properties)
      if properties.key? :start_page_number
        section_properties = SectionProperties.from_document(xml_node)
        section_properties.start_page_number = properties[:start_page_number]
      end
    end

    private
    def build_operations(fields)
      OperationConstruction.new(fields).operations
    end

    class Block < Struct.new(:start_field, :end_field)
      def self.enclosed_by(start_field, end_field)
        @blocks ||= [ImageBlock, RowBlock, ParagraphBlock]
        block_class = @blocks.detect { |klass| klass.encloses?(start_field, end_field) }
        block_class.new start_field, end_field
      end

      def process(context)
        replaced_node = Nokogiri::XML::Node.new('tmp', start_node.document)
        replaced_node.children = Nokogiri::XML::NodeSet.new(start_node.document, body.map(&:dup))
        Processor.process replaced_node, context
        replaced_node.children
      end

      def replace(content)
        content.each { |n| start_node.add_next_sibling n }

        body.each &:remove
        start_node.remove
        end_node.remove
      end

      def body
        return @body if defined?(@body)
        @body = []
        node = start_node
        while (node = node.next_element) && node != end_node
          @body << node
        end
        @body
      end

      def start_node
        @start_node ||= self.class.parent(start_field).first
      end

      def end_node
        @end_node ||= self.class.parent(end_field).first
      end

      def self.encloses?(start_field, end_field)
        parent(start_field).any? && parent(end_field).any?
      end
    end

    class RowBlock < Block
      def self.parent(node)
        node.ancestors './/w:tr'
      end

      def self.encloses?(start_field, end_field)
        if super
          parent(start_field) != parent(end_field)
        end
      end
    end

    class ParagraphBlock < Block
      def self.parent(node)
        node.ancestors './/w:p'
      end
    end

    class ImageBlock < ParagraphBlock
      def self.encloses?(start_field, end_field)
        start_field.expression =~ /^@/
      end

      def replace(content)
        pic_prop = self.class.parent(start_field).at_xpath('.//pic:cNvPr', 'pic' => PICTURE_NS_URI)
        pic_prop.attributes['name'].value = content.name
        blip = self.class.parent(start_field).at_xpath('.//a:blip', 'a' => MAIN_NS_URI)
        blip.attributes['embed'].value = "rId#{content.rid}"
        start_field.replace('')
        end_field.replace('')
      end
    end

    class OperationConstruction
      def initialize(fields)
        @fields = fields
        @operations = []
      end

      def operations
        while @fields.any?
          @operations << consume(true)
        end
        @operations.compact
      end

      def consume(allow_insertion)
        @field = @fields.shift
        return unless @field
        case @field.expression
        when /^=/
          if allow_insertion
            Statement::Insertion.new(Expression.parse(@field.expression[1..-1]), @field)
          end
        when /([^ ]+):each\(([^ ]+)\)/
          block = consume_block("#{$1}:endEach")
          Statement::Loop.new(Expression.parse($1), $2, block)
        when /([^ ]+):if\(([^)]+)\)/
          block = consume_block("#{$1}:endIf")
          Statement::Condition.new(Expression.parse($1), block, $2)
        when /([^ ]+):if/
          block = consume_block("#{$1}:endIf")
          Statement::Condition.new(Expression.parse($1), block)
        when /^@([^ ]+):start/
          block = consume_block("@#{$1}:end")
          Statement::Image.new(Expression.parse($1), block)
        end
      end

      def consume_block(end_expression)
        start_field = end_field = @field
        while end_field && end_field.expression != end_expression
          consume(false)
          end_field = @field
        end

        if end_field
          Block.enclosed_by start_field, end_field
        else
          raise TemplateError, "Could not find end field for «#{start_field.expression}». Was looking for «#{end_expression}»"
        end
      end
    end
  end
end
