require 'sinatra'
require 'roo'
require 'date'
require 'rack/protection'
require 'digest'
use Rack::Protection # for basic protection against simple attacks


# Prevents XSS and enforces HTTPS only
before do
  headers 'Content-Security-Policy' => "default-src 'self'; script-src 'self' 'unsafe-inline'"
  headers 'Strict-Transport-Security' => 'max-age=31536000; includeSubDomains'
  headers 'X-Content-Type-Options' => 'nosniff'
  headers 'X-Frame-Options' => 'DENY'
end

LEVELS = {
  6 => 2,
  12 => 3,
  16 => 4,
  24 => 5
}

def months_since(start_date, end_date = Date.today)
  return nil if start_date.nil? || start_date > end_date

  years = end_date.year - start_date.year
  months = end_date.month - start_date.month
  days = end_date.day - start_date.day

  total_months = years * 12 + months
  if days >= 15
    total_months += 1
    days = 0
  end
  total_months
end

def determine_current_level(months)
  return nil if months.nil?
  case months
  when 0...6 then 1
  when 6...12 then 2
  when 12...18 then 3
  when 18...24 then 4
  else 5
  end
end

def two_months_ago_range(today = Date.today)
  days_until_saturday = (6 - today.wday) % 7
  upcoming_saturday = today + days_until_saturday

  begin
    start_date = upcoming_saturday << 2 
  rescue ArgumentError
    start_date = Date.new(upcoming_saturday.year, upcoming_saturday.month - 2, -1)
  end

  end_date = start_date + 6
  (start_date..end_date)
end


def assessments_2_months_back(file)
  xlsx = Roo::Spreadsheet.open(file[:tempfile].path)
  sheet = xlsx.sheet(0)

  full_names = sheet.column(3)[1..]
  assessment_dates = sheet.column(12)[1..]

  date_range = two_months_ago_range(Date.today)

  parsed_dates = assessment_dates.map do |raw|
    begin
      case raw
      when Date then raw
      when Float, Integer then Date.new(1899, 12, 30) + raw.to_i
      when String then raw.strip.empty? ? nil : Date.strptime(raw.strip, '%m/%d/%Y')
      else nil
      end
    rescue
      nil
    end
  end

  full_names.zip(parsed_dates).map do |name, date|
    if date && date_range.cover?(date)
      "#{name}"
    end
  end.compact
end

def process_excel(file)
  xlsx = Roo::Spreadsheet.open(file[:tempfile].path)
  sheet = xlsx.sheet(0)

  first_names = sheet.column(3)[1..]
  last_names  = sheet.column(4)[1..]
  start_dates = sheet.column(11)[1..]

  full_names = first_names.zip(last_names).map { |f, l| "#{f} #{l}" }

  months_since_start = start_dates.map do |raw|
    begin
      parsed_date =
        case raw
        when Date
          raw
        when Float, Integer
          Date.new(1899, 12, 30) + raw.to_i
        when String
          raw.strip.empty? ? nil : Date.strptime(raw.strip, '%m/%d/%Y')
        else
          nil
        end
      parsed_date ? months_since(parsed_date) : nil
    rescue
      nil
    end
  end

  full_names.zip(months_since_start).map do |name, months|
    next unless LEVELS.key?(months)
    "#{name} levels up to level #{LEVELS[months]}"
  end.compact
end

def all_students_levels(file)
  xlsx = Roo::Spreadsheet.open(file[:tempfile].path)
  sheet = xlsx.sheet(0)

  first_names = sheet.column(3)[1..]
  last_names  = sheet.column(4)[1..]
  start_dates = sheet.column(11)[1..]

  full_names = first_names.zip(last_names).map { |f, l| "#{f} #{l}" }

  months_since_start = start_dates.map do |raw|
    begin
      parsed_date =
        case raw
        when Date then raw
        when Float, Integer then Date.new(1899, 12, 30) + raw.to_i
        when String then raw.strip.empty? ? nil : Date.strptime(raw.strip, '%m/%d/%Y')
        else nil
        end
      parsed_date ? months_since(parsed_date) : nil
    rescue
      nil
    end
  end

  full_names.zip(months_since_start).map do |name, months|
    level = determine_current_level(months)
    "#{name} is currently Level #{level}" if level
  end.compact
end


get '/' do
  <<-HTML
    <html>
      <head>
        <title>Student Level-Up Checker</title>
        <style>
          body {
            font-family: 'Segoe UI', sans-serif;
            background-color: #f5f7fa;
            margin: 0;
            padding: 2em;
          }
          .container {
            background-color: white;
            padding: 2em;
            max-width: 650px;
            margin: auto;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            border-radius: 12px;
          }
          h1 {
            text-align: center;
            color: #2c3e50;
            font-size: 1.75em;
            margin-bottom: 0.5em;
          }
          ul {
            padding-left: 1.25em;
            margin-bottom: 1.5em;
          }
          .note {
            background-color: #fff3cd;
            border: 1px solid #ffeeba;
            padding: 10px;
            margin-bottom: 20px;
            color: #856404;
            border-radius: 6px;
          }
          input[type="file"] {
            margin-bottom: 1em;
          }
          input[type="submit"] {
            background-color: #3498db;
            color: white;
            border: none;
            padding: 10px 20px;
            font-size: 1em;
            border-radius: 5px;
            cursor: pointer;
          }
          input[type="submit"]:hover {
            background-color: #2980b9;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h1>Student Level-Up Checker</h1>

          <div class="note">
            <strong>Note:</strong> If the page takes a few seconds to load, it’s just waking up — give it a moment!
          </div>

          <p><strong>How to use:</strong></p>
          <ul>
            <li>In Radius, download the enrollment report, with the date set to today</li>
            <li>Make sure to include only <em>in-person</em> and <em>currently enrolled</em> students</li>
            <li>Upload the Excel file below (if the enrollment report is not what is uploaded, you will receive strange/incorrect results)</li>
          </ul>

          <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" required><br>
            <input type="submit" name="action" value="calculate new level ups">
            <input type="submit" name="action" value="show all students current levels">
          </form>
          <a href="/assessments">Check Assessments</a>
        </div>
      </body>
    </html>
  HTML
end

get '/assessments' do
  <<-HTML
  <html>
    <head>
      <title>Upload Assessment File</title>
      <style>
        body { font-family: 'Segoe UI', sans-serif; background: #f5f7fa; padding: 2em; }
        .container {
          background: white; padding: 2em; max-width: 650px; margin: auto;
          box-shadow: 0 0 10px rgba(0,0,0,0.1); border-radius: 12px;
        }
        h1 { text-align: center; color: #2c3e50; }
        input[type="file"], input[type="submit"] {
          display: block; margin: 1em auto; padding: 10px;
        }
        a { text-align: center; display: block; margin-top: 2em; color: #3498db; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>Upload Assessment File</h1>
          <p><strong>How to use:</strong></p>
          <ul>
            <li>In Radius, under reports, press <em>Students</em> and download report</li>
            <li>Make sure to include only <em>in-person</em> and <em>currently enrolled</em> students</li>
            <li>Upload the Excel file below (if the student report is not what is uploaded, you will receive strange/incorrect results)</li>
          </ul>
        <form action="/assessments" method="post" enctype="multipart/form-data">
          <input type="file" name="file" required>
          <input type="submit" value="Check Assessments to Pull Forward">
        </form>
        <a href="/">← Go back to main level checker</a>
      </div>
    </body>
  </html>
  HTML
end

post '/upload' do
  if params[:file]
    filename = params[:file][:filename]

    # blocks malformed names or files not ending in .xls/.xlsx
    if filename =~ /[^a-zA-Z0-9_.\-\s]/ || !filename.end_with?('.xlsx', '.xls')
      return <<-HTML
        <p style="color:red;">Invalid file name or type. Please upload a clean Excel file (.xlsx or .xls).</p>
        <p><a href="/">Go back</a></p>
      HTML
    end

    # Prevents DoS via large file uploads
    if params[:file][:tempfile].size > 5 * 1024 * 1024
      return <<-HTML
        <p style="color:red;">File too large. Please upload a file under 5MB.</p>
        <p><a href="/">Go back</a></p>
      HTML
    end

    file_hash = Digest::SHA256.file(params[:file][:tempfile].path).hexdigest
    puts "[UPLOAD] #{filename} from #{request.ip} at #{Time.now} (SHA256: #{file_hash})"

    action = params[:action]

    result =
      case action
      when "calculate new level ups"
        process_excel(params[:file])
      when "show all students current levels"
        all_students_levels(params[:file])
      else
        ["Unknown action."]
      end
      <<-HTML
      <html>
        <head>
          <title>Level-Up Results</title>
          <style>
            body {
              font-family: 'Segoe UI', sans-serif;
              background-color: #f5f7fa;
              padding: 2em;
            }
            .container {
              background-color: white;
              padding: 2em;
              max-width: 650px;
              margin: auto;
              box-shadow: 0 0 10px rgba(0,0,0,0.1);
              border-radius: 12px;
            }
            h2 {
              color: #2c3e50;
              text-align: center;
              margin-bottom: 1em;
            }
            pre {
              background: #f4f4f4;
              padding: 1em;
              white-space: pre-wrap;
              word-break: break-word;
              border-radius: 6px;
              border: 1px solid #ddd;
            }
            a {
              display: inline-block;
              margin-top: 1.5em;
              text-decoration: none;
              color: #3498db;
              font-weight: bold;
              text-align: center;
            }
            a:hover {
              color: #1e70b8;
            }
            button#print-btn {
              margin-top: 1em;
              padding: 10px 15px;
              background-color: #2c3e50;
              color: white;
              border: none;
              border-radius: 5px;
              cursor: pointer;
              display: block;
              margin-left: auto;
              margin-right: auto;
            }
            @media print {
              body * {
                display: none !important;
              }
              #print-section {
                display: block !important;
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
              }
            }
          </style>
          <script>
            function printResults() {
              window.print();
            }
          </script>
        </head>
        <body>
          <div class="container">
            <div id="print-area">
              <h2>Level-Up Results</h2>
              <div id="print-section">
                <pre>#{result.empty? ? "No students ready to level up." : result.join("\n")}</pre>
              </div>
            </div>
            <button id="print-btn" onclick="printResults()">Print Results</button>
            <a href="/">← Upload another file</a>
          </div>
        </body>
      </html>
      HTML
  else
    "No file selected. <a href='/'>Try again</a>"
  end
end

post '/assessments' do
  if params[:file]
    filename = params[:file][:filename]

    if filename =~ /[^a-zA-Z0-9_.\-\s]/ || !filename.end_with?('.xlsx', '.xls')
      return "<p style='color:red;'>Invalid file type.</p><a href='/assessments'>Go back</a>"
    end

    if params[:file][:tempfile].size > 5 * 1024 * 1024
      return "<p style='color:red;'>File too large.</p><a href='/assessments'>Go back</a>"
    end

    file_hash = Digest::SHA256.file(params[:file][:tempfile].path).hexdigest
    puts "[ASSESSMENT UPLOAD] #{filename} from #{request.ip} at #{Time.now} (SHA256: #{file_hash})"

    results = assessments_2_months_back(params[:file])

    <<-HTML
    <html>
      <head>
        <title>Assessment Results</title>
        <style>
          body {
            font-family: 'Segoe UI', sans-serif;
            background-color: #f5f7fa;
            padding: 2em;
          }
          .container {
            background: white;
            padding: 2em;
            max-width: 650px;
            margin: auto;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            border-radius: 12px;
          }
          h2 {
            color: #2c3e50;
            text-align: center;
            margin-bottom: 1em;
          }
          pre {
            background: #f4f4f4;
            padding: 1em;
            white-space: pre-wrap;
            border-radius: 6px;
            border: 1px solid #ddd;
          }
          a {
            display: block;
            margin-top: 1.5em;
            text-align: center;
            color: #3498db;
            font-weight: bold;
            text-decoration: none;
          }
          a:hover {
            color: #1e70b8;
          }
          button#print-btn {
            margin-top: 1em;
            padding: 10px 15px;
            background-color: #2c3e50;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            display: block;
            margin-left: auto;
            margin-right: auto;
          }
          @media print {
            body * {
              display: none !important;
            }
            #print-section {
              display: block !important;
              position: absolute;
              top: 0;
              left: 0;
              width: 100%;
            }
          }
        </style>
        <script>
          function printResults() {
            window.print();
          }
        </script>
      </head>
      <body>
        <div class="container">
          <div id="print-area">
            <h2>Assessments to Pull</h2>
            <div id="print-section">
              <pre>#{result.empty? ? "No assessments to pull." : result.join("\n")}</pre>
            </div>
          </div>
          <button id="print-btn" onclick="printResults()">Print Results</button>
          <a href="/assessments">← Upload another file</a>
        </div>
      </body>
    </html>
    HTML


  else
    "No file selected. <a href='/assessments'>Try again</a>"
  end
end


