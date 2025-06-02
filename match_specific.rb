require 'roo'
require 'spreadsheet' 
require 'fileutils'

s1 = Roo::Excel.new("ApplicationC.xls")
s2 = Roo::Excel.new("DB2C.xls")

puts "ENTER Column 1 Name:"
col1 = gets.chomp!
puts "ENTER Parameter 1:"
param1 = gets.chomp!

puts "ENTER Column 2 Name:"
col2 = gets.chomp!
puts "ENTER Parameter 2:"
param2 = gets.chomp!

p "Start Running!!!"
c1=0
c2=0
r1=0
count = 0
txt1 = ""
txt2 = ""
(s1.first_column..s1.last_column).each do |column|
  if s1.cell(1,column).to_s.chomp == col1             # Finding column number of column Header
    c1 = column
  end  
  
  if s1.cell(1,column).to_s.chomp == col2             # Finding column number of column Header
    c2 = column
  end    
end

#puts "In 1st file, row number :"
(s1.first_row..s1.last_row).each do |row|  
  if s1.cell(row,c1).to_i.to_s.chomp == param1.to_s  &&  s1.cell(row,c2).to_s.chomp == param2.to_s      # Finding row number of given parameter
  # p r1 = row  
    txt1  =  txt1 + "row: #{row}: " + s1.row(row).to_s + "\n"                                   
   #p s1.row(row)
  end    
end
#puts "In 2nd file, row number :"
(s2.first_row..s2.last_row).each do |row| 
  if s2.cell(row,c1).to_i.to_s.chomp == param1.to_s  &&  s2.cell(row,c2).to_s.chomp == param2.to_s       # Finding row number of given parameter
 #  p r2 = row  
   txt2  =  txt2 + "row: #{row}: " + s2.row(row).to_s + "\n"
   count += 1                                   
   #p s1.row(row)
  end    
end
if count == 0
 # puts "No match found!!!!"
  txt2  =  txt2 + "No match found!!!!" + "\n"
end

# txt = ""


File.open( "first_file" + '.txt','w') do |s|
  s<<txt1
end

File.open( "second_file" + '.txt','w') do |s|
  s<<txt2
end

p "Done Scripting!!!"
