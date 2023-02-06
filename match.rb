require 'roo'
require 'spreadsheet' 
require 'fileutils'
require 'csv'
#require 'rubyXL'
#require 'xlsx_writer'

s1 = Roo::Excel.new("ordered_71_developer.xls")
s2 = Roo::Excel.new("ordered_71_staging.xls")

# doc = XlsxWriter.new
# sheet1 = doc.add_sheet("New")
# sheet1.add_row(["Row Number", "Column Number", "Source 1 Content", "Source 2 Content"])

p "Start Running!!!"
count = 0
txt = ""
name=""
individual_name= ""
CSV.open("error_log.csv", "wb") do |csv|
    csv << ["Column Name","Row Number", "Column Number", "Source 1 Content", "Source 2 Content"]
  #csv << ["fox", "1", "$90.00"]
#end

(s1.first_row..s1.last_row).each do |row|
  puts row
  ([*s1.first_column..s1.last_column]-[7,19,23,26,30]).each do |column|
  #(s1.first_column..s1.last_column).each do |column|
    #p column
    #txt = ""
  #  if column != 12 && column != 13 && column != 16 && column != 17
      if s1.cell(row,column).to_s.chomp != s2.cell(row,column).to_s.chomp
        puts "..........********............"
        puts s1.cell(row,column).to_s.chomp
        puts s2.cell(row,column).to_s.chomp
        puts "..........********............"
        puts "Mismatch FOUND!!!"
        puts "row: #{row}, column: #{column} "
        puts "Mismatch:  #{s1.cell(row,column).to_s}"
        txt  =  txt + "row: #{row}, column: #{column} " + "\n"
        txt  =  txt  +  "Mismatch: " + s1.cell(row,column).to_s + "\n"

        csv << [s1.cell(1,column).to_s, "#{row}", "#{column}", s1.cell(row,column).to_s, s2.cell(row,column).to_s]
       # sheet1.add_row(["#{row}", "#{column}", s1.cell(row,column).to_s, s2.cell(row,column).to_s])

          if individual_name != s1.cell(row,3).to_s.chomp
            individual_name = s1.cell(row,3).to_s
            name = name + "#{individual_name}" + "\n"
            count += 1
          end
   #   end
      end
     end
  end
end
# ::FileUtils.mv doc.path, 'myfile.xlsx'
# doc.cleanup
File.open( "error_log" + '.txt','w') do |s|
  s<<txt
end

File.open( "error_individual" + '.txt','w') do |s|
  s<<name
end



p "Total Mismatch Count: #{count}"
p "Done Scripting!!!"
