require "google_drive"

Session = GoogleDrive.login(ENV['GOOGLE_USER'], ENV['GOOGLE_PASS'])

class SpreadsheetWriter 
  attr_reader :worksheet, :row, :rows, :spreadsheet
 

  def initialize(spreadsheet)
    @spreadsheet = spreadsheet
    @worksheet = Session.spreadsheet_by_url(spreadsheet).worksheets[0]
    @rows = @worksheet.rows.count
    @row = rows + 1
  end

  def write_to_spreadsheet(input, column = 1)
    if input.kind_of?(Array)
      array_instance = ArraySpreadsheetWriter.new(spreadsheet)
      return array_instance.write_to_spreadsheet(input)
    elsif input.kind_of?(Hash)
      hash_instance = HashSpreadsheetWriter.new(spreadsheet)
      return hash_instance.write_to_spreadsheet(input)
    else
      populate_row_with(input, column) 
      next_row
    end
  end

  def populate_row_with(input, column)
    worksheet[row, column] = input
  end

  def next_row
    @row += 1
    refresh_worksheet
  end

  def refresh_worksheet
    worksheet.synchronize
    worksheet.reload
  end

  def write_array_down_column(array, column)
    array.each do |item|
      populate_row_with(item, column)
      next_row
    end
  end

  def write_array_across_row(array)
    column = 1
    array.each do |item|
      populate_row_with(item, column)
      column = column + 1
      refresh_worksheet
    end
  end
end


class HashSpreadsheetWriter < SpreadsheetWriter

  def write_to_spreadsheet(input)
    input.each do |key, value|
      key_column = 1
      value_column = 2
      populate_row_with(key, key_column)
      if value.kind_of?(Array)
        write_array_down_column(value, value_column)
      else
        populate_row_with(value, value_column)
        next_row
      end
    end
  end
end


class ArraySpreadsheetWriter < SpreadsheetWriter

  def write_to_spreadsheet(array)
    array.each do |entry|
      if entry.kind_of?(Array) 
        write_array_across_row(entry)
      else
        populate_row_with(entry, column)
      end
      next_row
    end
  end
end





