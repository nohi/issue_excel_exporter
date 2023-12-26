require 'redmine'
require_dependency 'query'

Redmine::Plugin.register :issue_excel_exporter do
  name 'Issue Excel Exporter plugin'
  author 'nohi'
  description 'This plugin outputs Issues as xlsx files, indenting the Title based on the parent-child relationship of the Issue.'
  version '0.0.1'
end

# plugin view hooks
Dir[File.dirname(__FILE__) + '/lib/issue_excel_exporter/*.rb'].each do |file|
  require_dependency file
end