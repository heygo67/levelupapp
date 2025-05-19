require 'sinatra'
require 'roo'
require 'date'

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
  total_months += 1 if days > 0
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
            font-family: sans-serif;
            padding: 2em;
            max-width: 600px;
            margin: auto;
          }
          h1 { font-size: 1.5em; }
          pre { background: #f4f4f4; padding: 1em; }
        </style>
      </head>
      <body>
        <h1>Student Level-Up Checker</h1>
        <p><strong>How to use:</strong></p>
        <ul>
          <li>In Radius, download the enrollment report, with the date set to today</li>
          <li>Make sure to include only <em>in-person</em> and <em>currently enrolled</em> students</li>
          <li>Upload the Excel file below</li>
        </ul>

        <form action="/upload" method="post" enctype="multipart/form-data">
          <input type="file" name="file" required>
          <input type="submit" value="Upload">
        </form>
      </body>
    </html>
  HTML
end

post '/upload' do
  if params[:file]
    result = process_excel(params[:file])
    <<-HTML
      <html>
        <head>
          <title>Results</title>
        </head>
        <body>
          <h2>Level-Up Results</h2>
          <pre>#{result.empty? ? "No students ready to level up." : result.join("\n")}</pre>
          <p><a href="/">Upload another file</a></p>
        </body>
      </html>
    HTML
  else
    "No file selected. <a href='/'>Try again</a>"
  end
end
