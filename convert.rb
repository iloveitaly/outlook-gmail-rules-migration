#!/usr/bin/env ruby

require 'rubygems'
require 'fastercsv'

# rule data is formatted in a "email, email box path" format
rule_data = FasterCSV.table("sample_rules.csv", :force_quotes => true)

# example gmail rules:
# <entry>
#   <category term='filter'></category>
#   <title>Mail Filter</title>
#   <id>tag:mail.google.com,2008:filter:1326039218652</id>
#   <updated>2012-01-12T19:19:49Z</updated>
#   <content></content>
#   <apps:property name='from' value='donotreply@airbrake.io'/>
#   <apps:property name='label' value='Services'/>
#   <apps:property name='shouldArchive' value='true'/>
# </entry>

gmail_rules = ""
template = %Q(
<entry>
  <category term='filter'></category>
  <title>%s</title>
  <apps:property name='from' value='%s'/>
  <apps:property name='label' value='%s'/>
  <apps:property name='shouldArchive' value='true'/>
</entry>)

rule_data.each do |row|
  gmail_rules += template % [row[1], row[0], row[1]]
end

puts "<?xml version='1.0' encoding='UTF-8'?><feed xmlns='http://www.w3.org/2005/Atom' xmlns:apps='http://schemas.google.com/apps/2006'>" + gmail_rules + "</feed>"