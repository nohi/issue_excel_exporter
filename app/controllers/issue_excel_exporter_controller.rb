require 'rubyXL'

class IssueExcelExporterController < ApplicationController
  include QueriesHelper

  def export
    @project = Project.find(params[:project_id])
    
    use_session = false
    retrieve_default_query(use_session)
    retrieve_query(IssueQuery, use_session)

    issues = @query.issues(:limit => Setting.issues_export_limit.to_i)
    workbook = RubyXL::Workbook.new
    worksheet = workbook[0]

    # チケットの階層構造に基づくデータの追加
    sorted_issues = sort_issues(issues)

    # チケットデータをworksheetに書き込む
    write_worksheet(worksheet, sorted_issues, @query)

    # xlsxを出力
    send_data workbook.stream.string, filename: 'issues.xlsx', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  end

  private

  def retrieve_default_query(use_session)
    return if params[:query_id].present?
    return if api_request?
    return if params[:set_filter]

    if params[:without_default].present?
      params[:set_filter] = 1
      return
    end
    if !params[:set_filter] && use_session && session[:issue_query]
      query_id, project_id = session[:issue_query].values_at(:id, :project_id)
      return if IssueQuery.where(id: query_id).exists? && project_id == @project&.id
    end
    if default_query = IssueQuery.default(project: @project)
      params[:query_id] = default_query.id
    end
  end

  # チケットの親子構造に基づきデータをソートする
  # ex:
  # - チケット
  #   - 子チケット
  #    - 孫チケット
  def sort_issues(issues)
    # チケットを親IDによってグループ化
    issue_groups = issues.group_by(&:parent_id)
  
    # トップレベルのチケット（親IDがないもの）から再帰的にソート
    sorted = sort_issue_group(issue_groups, nil)
    sorted
  end
  
  def sort_issue_group(issue_groups, parent_id, level = 0)
    sorted = []
    if issue_groups[parent_id]
      issue_groups[parent_id].each do |issue|
        # 各チケットとそのレベルを追加
        sorted << { issue: issue, level: level }
        # 子チケットを再帰的に追加
        sorted.concat(sort_issue_group(issue_groups, issue.id, level + 1))
      end
    end
    sorted
  end

  # チケットデータをworksheetに書き込む
  def write_worksheet(worksheet, sorted_issues, query)
    # ヘッダー行の追加
    query.columns.each_with_index do |column, index|
      worksheet.add_cell(0, index, column.caption)
    end
  
    # 題名に相当するカラムのインデックスを見つける
    title_column = query.columns.find { |c| c.name == :subject }
  
    # 各チケットのデータ行を追加
    sorted_issues.each_with_index do |item, row|
        issue = item[:issue]
        level = item[:level]
        query.columns.each_with_index do |column, col|
          value = value_for_column(issue, column)
    
          # 題名のカラムのみにインデントを適用
          value = ' ' * (2 * level) + value if column.name == title_column.name
    
          worksheet.add_cell(row + 1, col, value)
        end
      end
  end
  
  # 特定のカラムに対するチケットの値を取得するメソッド
  def value_for_column(issue, column)
    value = column.value(issue)
  
    # 値がオブジェクトの場合は、名前またはタイトルを返す
    if value.is_a?(ActiveRecord::Base) && value.respond_to?(:name)
      value.name
    elsif value.is_a?(ActiveRecord::Base) && value.respond_to?(:title)
      value.title
    else
      value.to_s
    end
  end
end