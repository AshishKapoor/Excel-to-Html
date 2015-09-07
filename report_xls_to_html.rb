=begin
 *****************************************************************************************
 File: reporthtml.rb
 *   Description: This file is for rendering the generated report in xls to HTML page.

 *   Revision History:
 *   Date         Name              Reference    Purpose
 *   17/Aug/2015  Ashish Kapoor     N/A          Initial Creation
 *   Copyright : TBD
 *****************************************************************************************
=end

require 'spreadsheet'

File.open('data.csv', 'w') do |file|
book_descriptor = Spreadsheet.open('TestReport_2015_8_17_18_8_46.xls')
sheet_video = book_descriptor.worksheet('Test_Report_Video') # worksheet name
sheet_video.each do |row|
    break if row[0].nil? # if in case the first cell is empty
    file.puts row.join(', ') # saves the output in the data.csv file
    end
end
System.getProperty("os.name")
