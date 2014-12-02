module Sablon
  class Template
    def initialize(path)
      @path = path
    end

    # Same as +render_to_string+ but writes the processed template to +output_path+.
    def render_to_file(output_path, context, images = [], properties = {})
      File.open(output_path, 'w') do |f|
        f.write render_to_string(context, images, properties)
      end
    end

    # Process the template. The +context+ hash will be available in the template.
    def render_to_string(context, images = [], properties = {})
      render(context, images, properties).string
    end

    private
    def render(context, images, properties)
      Zip::OutputStream.write_buffer do |out|
        images.each do |image|
          if image.name =~ /^(.*)\.jpg$/
            image.name = "#{$1}.jpeg"
          end
          out.put_next_entry(File.join('word', 'media', image.name))
          out.write(image.data)
        end
        Zip::File.open(@path).each do |entry|
          entry_name = entry.name
          out.put_next_entry(entry_name)
          content = entry.get_input_stream.read
          if entry_name == 'word/document.xml'
            xml_node = Processor.process(Nokogiri::XML(content), context, properties)
            Processor.remove_final_blank_page xml_node
            out.write(xml_node.to_xml)
          elsif entry_name =~ /word\/header\d*\.xml/ || entry_name =~ /word\/footer\d*\.xml/
            out.write(Processor.process(Nokogiri::XML(content), context).to_xml)
          elsif entry_name == 'word/_rels/document.xml.rels' && !images.empty?
            out.write(Processor.process_rels(Nokogiri::XML(content), images).to_xml)
          else
            out.write(content)
          end
        end
      end
    end
  end
end
