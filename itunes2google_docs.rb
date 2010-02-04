#!/usr/local/bin/ruby

require 'rubygems'
require 'mechanize'
require 'gdata'
require 'optparse'
require 'xmlsimple'
require 'ruby-debug'
require './itunes2google_docs_defaults.rb'

# Parse command line options
def parse_command_line
  # This hash will hold all of the options parsed from the command-line by OptionParser.
  options = {}
  optparse = OptionParser.new do |opts|
    # Set a banner, displayed at the top of the help screen.
    opts.banner = "Usage: #{File.basename($0)} [options]\n"
    opts.banner += "Reads Daily and Weekly summary files from iTunes Connect, adds to Google Docs folder and adds to Google Spreadsheet\n"
    opts.banner += "Default is --daily\n"
    opts.banner += "Folder ids are found in browsers at http://docs.google.com/?tab=co&pli=1#folders/folder.0.FOLDERID\n"
    opts.banner += "Folder URIS for uploading are of form http://docs.google.com/feeds/folders/private/full/folder%3AFOLDERID\n"
    opts.banner += "Construct the worksheets feed URI by http://spreadsheets.google.com/feeds/worksheets/SPREADSHEETID/private/full\n"
    opts.banner += "Worksheet URIs look something like http://spreadsheets.google.com/feeds/list/SPREADSHEETID/od6/private/full\n"
    opts.banner += "\nOptions are:\n"
 
    options[:itunes_username] = nil
    opts.on('-u', '--itunes_username', 'itunes username VALUE') do |value|
      options[:itunes_username] = value
    end
   
     options[:itunes_password] = nil
    opts.on('-p', '--itunes_password', 'itunes password VALUE') do |value|
      options[:itunes_password] = value
    end
    
       options[:google_username] = nil
    opts.on('-x', '--google_username', 'google username VALUE') do |value|
      options[:google_username] = value
    end
   
     options[:google_password] = nil
    opts.on('-y', '--google_password', 'google password VALUE') do |value|
      options[:google_password] = value
    end
    
    options[:google_docs_daily_folder] = nil
    opts.on('-a', '--google_docs_daily_folder', 'google docs daily folder VALUE') do |value|
      options[:google_docs_daily_folder] = value
    end
    
    options[:google_docs_weekly_folder] = nil
    opts.on('-b', '--google_docs_weekly_folder', 'google docs weekly folder VALUE') do |value|
      options[:google_docs_weekly_folder] = value
    end

   options[:google_docs_financials_folder] = nil
    opts.on('-c', '--google_docs_financials_folder', 'google docs financials folder VALUE') do |value|
      options[:google_docs_financials_folder] = value
    end

   options[:google_spreadsheet_daily_worksheet_url] = nil
    opts.on('-j', '--google_spreadsheet_daily_worksheet_url', 'google docs spreadsheet daily worksheet url VALUE') do |value|
      options[:google_spreadsheet_daily_worksheet_url] = value
    end

   options[:google_spreadsheet_weekly_worksheet_url] = nil
    opts.on('-e', '--google_spreadsheet_weekly_worksheet_url', 'google docs spreadsheet weekly worksheet url VALUE') do |value|
      options[:google_spreadsheet_weekly_worksheet_url] = value
    end

    options[:google_spreadsheet_financials_worksheet_url] = nil
    opts.on('-k', '--google_spreadsheet_financials_worksheet_url', 'google docs spreadsheet financials worksheet url VALUE') do |value|
      options[:google_spreadsheet_financials_worksheet_url] = value
    end

      options[:verbose] = false
      opts.on( '-v', '--verbose', 'Output more information' ) do
        options[:verbose] = true
      end
      
      options[:weekly] = false
      opts.on( '-w', '--weekly', 'Do Weekly summary' ) do
        options[:weekly] = true
      end

      options[:daily] = false
      opts.on( '-d', '--daily', 'Do Daily summary' ) do
        options[:daily] = true
      end
      
       options[:financials] = false
      opts.on( '-$', '--financials', 'Do financials' ) do
        options[:financials] = true
      end
     
      options[:read_itunes] = true
      opts.on( '-t', '--no_itunes', 'Do not read from iTunes' ) do
        options[:read_itunes] = false
      end

    options[:write_google_docs] = true
      opts.on( '-g', '--no_google_docs', 'Do not output to Google Docs' ) do
        options[:write_google_docs] = false
      end

    options[:write_google_spreadsheet] = true
      opts.on( '-s', '--no_google_spreadsheet', 'Do not output to Google Spreadsheet' ) do
        options[:write_google_spreadsheet] = false
      end

      options[:logfile] = nil
      opts.on( '-l', '--log FILE', 'Write log to FILE' ) do|file|
         options[:logfile] = file
        end
        
    options[:list_spreadsheets] = false
      opts.on( '-m', '--list_spreadsheets', 'list spreadsheets in google docs' ) do
         options[:list_spreadsheets] = true
    end
    
    options[:list_worksheets] = nil
      opts.on( '-n', '--list_worksheets VALUE', 'list worksheets in spreadsheet VALUE' ) do |value|
         options[:list_worksheets] = value
    end


      options[:filename] = nil
            opts.on( '-f', '--filename VALUE', 'Use file name VALUE') do |value|
        options[:filename] = value
      end
 
    # This displays the help screen, all programs are
    # assumed to have this option.
    opts.on( '-h', '--help', 'Display this screen' ) do
      puts opts
      exit
    end
  end
 
  # Parse the command-line. Remember there are two forms of the parse method. The 'parse' method simply parses
  # ARGV, while the 'parse!' method parses ARGV and removes any options found there, as well as any parameters for
  # the options. What's left is the list of files to resize.
  optparse.parse!
 
 return options
  # Exit if no file names are given on the command line
#  filenames = ARGV
#  if filenames.empty?
#    puts optparse.help
#    exit
#  end
  
#  return filenames, options
end

def apply_defaults(options)
            
    Defaults.each {|key,value|
        options[key] = value if options[key].nil?
    }
            
    pp Defaults if options[:verbose]
    pp options if options[:verbose]
end

def fetch_itunes_connect_document(options)

    puts ("Connecting to iTunes Connect") if options[:verbose]
    agent = WWW::Mechanize.new
    page = agent.get('https://itunesconnect.apple.com/')
    form = page.form('appleConnectForm')
    form.theAccountName = options[:itunes_username]
    form.theAccountPW = options[:itunes_password]
    page = agent.submit(form, form.buttons.first)
    
    # The page result sometimes has a continue button when there's a new licensing agreement
    mainForm = page.form('mainForm')
    if mainForm != nil
        page = agent.submit(mainForm, mainForm.buttons.first)
    end
        
    if options[:financials]
        page = agent.click page.link_with(:text => 'Finance Reports')
        links = page.links
        options[:filenames] = []
        links.each { |link|
            if ( link.text =~ /\n *([^ ]*)\n/) == 0
                page4 = link.click
                f = open($1, "w")
                options[:filenames].push($1)
                puts( "Writing to #{$1}") if options[:verbose]
                f.write(page4.body)
                f.close
                end
        }
     end
    
    if options[:weekly] || options[:daily]
        page = agent.click page.link_with(:text => 'Sales and Trends')
        page2 = agent.get('https://itts.apple.com/cgi-bin/WebObjects/Piano.woa')
        form2 = page2.form('frmVendorPage')
        form2.hiddenSubmitTypeName = 'Summary'
        if options[:daily]
            form2.field_with(:name => '11.11').options[3].select
            form2.hiddenDayOrWeekSelection = 'Daily'
            page3 = agent.submit(form2) #, form2.buttons[2])
            form3 = page3.form('frmVendorPage')
            form3.field_with(:name => '11.13.1').options[0].select
            form3.hiddenSubmitTypeName = 'Download'
            page4 = agent.submit(form3) #, form3.buttons[2])
            options[:filename] = page4.filename.sub(/\.gz/, "") if options[:filename].nil?
            f = open(options[:filename], "w")
            puts( "Writing to #{options[:filename]}") if options[:verbose]
            f.write(page4.body)
            f.close
        end
        if options[:weekly]
            form2.field_with(:name => '11.11').options[2].select
            form2.hiddenDayOrWeekSelection = 'Weekly'
            page3 = agent.submit(form2) #, form2.buttons[2])
            form3 = page3.form('frmVendorPage')
            form3.field_with(:name => '11.15.1').options[0].select
            form3.hiddenSubmitTypeName = 'Download'
            page4 = agent.submit(form3) #, form3.buttons[2])
            options[:filename] = page4.filename.sub(/\.gz/, "") if options[:filename].nil?
            f = open(options[:filename], "w")
            puts( "Writing to #{options[:filename]}") if options[:verbose]

            f.write(page4.body)
            f.close
        end
    end


    return true
end

def upload_to_google_docs(options)
    puts "Connecting to Google Docs" if options[:verbose]
    client = GData::Client::DocList.new
    client.clientlogin(options[:google_username], options[:google_password])
    if options[:daily]
        puts "Uploading Daily from #{options[:filename]}" if options[:verbose]
        response  = client.post_file(options[:google_docs_daily_folder], options[:filename], "text/plain") 
    end
    if options[:weekly]
        puts "Uploading Weekly from #{options[:filename]}" if options[:verbose]
        response  = client.post_file(options[:google_docs_weekly_folder], options[:filename], "text/plain")
    end
    if options[:financials]
        options[:filenames].each {|filename|
            puts "Uploading financials from #{filename}" if options[:verbose]
            response  = client.post_file(options[:google_docs_financials_folder], filename, "text/plain")
        }
    end

end

def upload_to_google_spreadsheet(options, filename, worksheet_url)
    puts "Connecting to Google Spreadsheet" if options[:verbose]
    client2 = GData::Client::Spreadsheets.new
    client2.clientlogin(options[:google_username], options[:google_password])
    
    File.open(filename, 'r') do |input_file|
        puts "Parsing file #{filename}" if options[:verbose]
      
        # First line: headings
        column_headers = []
        fields = []
        first_line = input_file.gets
        field_names = first_line.split("\t")
        field_names.each do |field_name|
          field_name.strip!
          field_name.gsub!(/ /, "")
          field_name.gsub!(/\//, "")
          field_name.downcase!
          fields << field_name
          
          # Store column headers
          column_headers << field_name
        end
      
        # Other lines: data
        data_rows = []
        total_amount_currency = ""
                puts( "Posting file #{filename}.atom to #{worksheet_url}") if options[:verbose]
        while (line = input_file.gets)
            new_row = '<atom:entry xmlns:atom="http://www.w3.org/2005/Atom" xmlns:gsx="http://schemas.google.com/spreadsheets/2006/extended">'

          # Skip if line is empty or contains only whitespace characters (such as a line of empty tabs)
          next if line =~ /^\s*$/ || line =~ /Total_Amount/
          
          # Detect Totals row at the bottom of the table.
          # Current format: Total_Amount:32.68
          # Old format (pre-Feb 2009): \t\t\t\t\t\tTotal\t32.68 AUD\t...
          break if (line =~ /^Total_Amount/) || (line =~ /^\s*Total\t/)
          
          # Else: we have a data row => start processing
          row = []
          field_values = line.split("\t")
          field_values.each_with_index do |field_value, index|
            field_value.strip!
            field_data = { :value => field_value }
            current_field = fields[index]
            new_row += "<gsx:#{current_field}>#{field_value}</gsx:#{current_field}>"
          end
        new_row += '</atom:entry>'
        f = open(filename+".atom", "w")
        f.write(new_row)
        f.close
        response = client2.post_file(worksheet_url, filename+".atom", "application/atom+xml")
        end
      end
end

def list_spreadsheets(options)
    client = GData::Client::Spreadsheets.new
    client.clientlogin(options[:google_username], options[:google_password])
    response = client.get("http://spreadsheets.google.com/feeds/spreadsheets/private/full")
    spreadsheet_list =  XmlSimple.xml_in(response.body, 'KeyAttr' => 'name')
    pp spreadsheet_list
end

def list_worksheets(options)
    client = GData::Client::Spreadsheets.new
    client.clientlogin(options[:google_username], options[:google_password])
    response = client.get(options[:list_worksheets])
    worksheet_list =  XmlSimple.xml_in(response.body, 'KeyAttr' => 'name')
    pp worksheet_list
end


options = parse_command_line()
apply_defaults(options)
  if(!options[:daily] && !options[:weekly] && !options[:financials] && !options[:list_spreadsheets] && options[:list_worksheets].nil? )
    puts ("Daily, weekly, list spreadsheets or list worksheets required")
    exit
  end

$stdout = File.new(options[:logfile], "w") if options[:logfile]
fetch_itunes_connect_document(options) if options[:read_itunes]
list_spreadsheets(options) if options[:list_spreadsheets]
list_worksheets(options) if options[:list_worksheets]
upload_to_google_docs(options) if options[:write_google_docs]
upload_to_google_spreadsheet(options, options[:filename], options[:google_spreadsheet_daily_worksheet_url]) if options[:write_google_spreadsheet] && options[:daily]
upload_to_google_spreadsheet(options, options[:filename], options[:google_spreadsheet_weekly_worksheet_url]) if options[:write_google_spreadsheet] && options[:weekly]
if options[:write_google_spreadsheet] && options[:financials]
    options[:filenames].each {|filename|
        upload_to_google_spreadsheet(options, filename, options[:google_spreadsheet_financials_worksheet_url])
    }
end






