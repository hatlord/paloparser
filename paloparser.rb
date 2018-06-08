#!/usr/bin/env ruby
##paloparser is a PaloAlto Firewall XML to CSV/Excel convertor.
require 'nokogiri'
require 'writeexcel'
require 'pp'


@palo_config = Nokogiri::XML(File.read(ARGV[0]))
@filename    = File.basename ARGV[0], ".xml"
@filename = @filename.gsub(" ", "_")


def parse_xml
  puts "Parsing XML...."
  @rule_array = []
  puts "Palo Alto Version: #{@palo_config.xpath('//config/@version').text}"
  @palo_config.xpath('//config/devices/entry/vsys/entry').each do |entry|
    @entry = {}
    @entry[:fwname] = entry.xpath('@name').text
    @entry[:rules]  = []
      entry.xpath('./rulebase/security/rules/entry').each do |rule|
      rules = {}
      rules[:fwname]        = @entry[:fwname]
      rules[:rulename]      = rule.xpath('@name').text
      rules[:frominterface] = rule.xpath('./to/member').map(&:text).join("\r")
      rules[:tointerface]   = rule.xpath('./from/member').map(&:text).join("\r")
      rules[:source]        = rule.xpath('./source/member').map(&:text).join("\r")
      rules[:destination]   = rule.xpath('./destination/member').map(&:text).join("\r")
      rules[:sourceuser]    = rule.xpath('./source-user/member').map(&:text).join("\r")
      rules[:category]      = rule.xpath('./category/member').map(&:text).join("\r")
      rules[:application]   = rule.xpath('./application/member').map(&:text).join("\r")
      rules[:service]       = rule.xpath('./service/member').map(&:text).join("\r")
      rules[:hip_profiles]  = rule.xpath('./hip-profiles/member').map(&:text).join("\r")
      rules[:action]        = rule.xpath('./action').map(&:text).join("\r")
      rules[:description]   = rule.xpath('./description').map(&:text).join("\r")

      @entry[:rules] << rules
    end
    @rule_array << @entry
  end
  @rule_array
end

def create_excel_file
  @xls_file  = "#{@filename}_Rules_EXCEL_#{Time.now.strftime("%d%b%Y_%H%M%S")}.xls"
  @rules_xls = WriteExcel.new("#{@xls_file}")
end


def create_sheets
  puts "Creating Excel File..."
  rule_array = []
  headers    = ['VSYS', 'Name', 'From Interface', 'To Interface', 'Source', 'Destination', 'User', 'Category', 'Application', 'Service', 'HIP-Profiles', 'Action', 'Description']
  excel = @rules_xls.add_worksheet("RULES")
  excel.write_col('A1', [headers])
  parse_xml.each do |outer|
    outer[:rules].each do |rules|
      rule_array << rules.values
    end
  end
  excel.write_col('A2', rule_array)
  @rules_xls.close
  puts "...Done"
end

create_excel_file
create_sheets
