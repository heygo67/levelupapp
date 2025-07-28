require 'sinatra'
require 'roo'
require 'date'
require 'rack/protection'
require 'digest'
use Rack::Protection # for basic protection against simple attacks


# Prevents XSS and enforces HTTPS only
before do
  headers 'Content-Security-Policy' => "default-src 'self'"
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
  total_months += 1 if days >= 15
  total_months
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
            <li>Upload the Excel file below</li>
          </ul>

          <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" required><br>
            <input type="submit" value="Upload">
          </form>
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

    result = process_excel(params[:file])
    <<-HTML
      <html>
        <head>
          <title>Results</title>
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
            }
            a:hover {
              color: #1e70b8;
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h2>Level-Up Results</h2>
            <pre>#{result.empty? ? "No students ready to level up." : result.join("\n")}</pre>
            <a href="/">← Upload another file</a>
          </div>
        </body>
      </html>
    HTML
  else
    "No file selected. <a href='/'>Try again</a>"
  end
end

