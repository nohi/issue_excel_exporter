# plugins/issue_excel_exporter/config/routes.rb
Rails.application.routes.draw do
    get 'issue_excel_exporter/export', to: 'issue_excel_exporter#export'
    get '/projects/:project_id/issue_excel_exporter/export', to: 'issue_excel_exporter#export'
end
  