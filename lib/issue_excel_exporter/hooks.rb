module IssueExcelExporter
    class Hooks < Redmine::Hook::ViewListener
      render_on :view_issues_index_bottom, :partial => "export_button"
    end
end
