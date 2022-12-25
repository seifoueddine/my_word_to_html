# frozen_string_literal: true

require_relative "my_word_to_html_gem/version"

module MyWordToHtmlGem
  require "rubyzip"

  class Converter
    def initialize(file_path)
      @file_path = file_path
    end

    def convert
      zip_file = RubyZip::File.open(@file_path)
      word_document = zip_file.find_entry("word/document.xml")

      # Read the contents of the word document
      doc_content = word_document.get_input_stream.read
      # Replace all instances of "w:p" with "p" to convert to HTML
      doc_content.gsub!("w:p", "p")

      # Return the converted HTML
      "<html><body>#{doc_content}</body></html>"
    end
  end
end
