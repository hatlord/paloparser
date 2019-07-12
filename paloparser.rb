#!/usr/bin/env ruby
##paloparser is a PaloAlto Firewall XML to CSV/Excel convertor.
$VERBOSE = nil
require 'nokogiri'
require 'axlsx'
require 'pp'


@palo_config = Nokogiri::XML(File.read(ARGV[0]))
@filename    = File.basename ARGV[0], ".xml"
@filename = @filename.gsub(" ", "_")


def parse_xml(xml_section, rule_type)
  @rule_array = []
  @palo_config.xpath(xml_section).each do |entry|
    @entry = {}
    @entry[:fwname] = entry.xpath('@name').text
    @entry[:rules]  = []
      entry.xpath(rule_type).each do |rule|
      rules = {}
      rules[:fwname]        = @entry[:fwname]
      rules[:rulename]      = rule.xpath('@name').text
      rules[:frominterface] = rule.xpath('./from/member').map(&:text).join("\r")
      rules[:tointerface]   = rule.xpath('./to/member').map(&:text).join("\r")
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
  @excel_file = "Palo_Rules_Excel_#{Time.now.strftime("%d%b%Y_%H%M%S")}.xlsx"
  @p = Axlsx::Package.new
  @wb = @p.workbook
end

def create_excel_data
  create_excel_file
  rule_types = {rulebase_security: ['//config/devices/entry/vsys/entry', './rulebase/security/rules/entry'], pre_rulebase_sec: ['//config/devices/entry/device-group/entry', './pre-rulebase/security/rules/entry'], post_rulebase_sec: ['//config/devices/entry/device-group/entry', './post-rulebase/security/rules/entry'], pre_rulebase_decrypt: ['//config/devices/entry/device-group/entry', './pre-rulebase/decryption/rules/entry'], post_rulebase_decrypt: ['//config/devices/entry/device-group/entry', './post-rulebase/decryption/rules/entry']}
  headers    = ['VSYS/DeviceGroup', 'Name', 'From Interface', 'To Interface', 'Source', 'Destination', 'User', 'Category', 'Application', 'Service', 'HIP-Profiles', 'Action', 'Description']
  rule_types.each do |key, value|
    @wb.add_worksheet(:name => key.to_s) do |sheet|
      sheet.add_row(headers)
      parse_xml(value[0], value[1]).each do |outer|
        outer[:rules].each do |rules|
          sheet.add_row rules.values.to_a
        end
      end
    end
    @p.serialize @excel_file
  end
  puts "Rules written to #{@excel_file}"
end

create_excel_data
