require 'roo'
require 'spreadsheet'
require 'fileutils'
require 'csv'

s1 = Roo::Excel.new("comp.xls")
match=""
(2..s1.last_row).each do |row1|
  puts row1
  match=""
  (2..54).each do |row2|
    if s1.cell(row1,1).to_s.chomp == s1.cell(row2,2).to_s.chomp
      match = "FOUND!!"
      break
    end

  end

  if match != "FOUND!!"
    puts s1.cell(row1,1).to_s.chomp
  end

  end