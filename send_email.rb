require 'mail'
require 'dotenv'
require 'creek'

# Load credentials from .env
Dotenv.load

# --- CONFIGURATION ---
SMTP_CONFIG = {
  address:              '127.0.0.1',
  port:                 1025,
  user_name:            ENV['PROTON_BRIDGE_USER'],
  password:             ENV['PROTON_BRIDGE_PASS'],
  authentication:       'plain',
  enable_starttls_auto: true,
  # WARNING: 'none' is insecure. This is for local Proton Bridge.
  # For other SMTP servers, use :peer or :client_once
  openssl_verify_mode:  'none'
}

DOC_LINK      = "https://bit.ly/3YaNO35"
PHOTO_FILE    = "header_photo.jpg"
PDF_FILE      = "Christmas Letter 2025.pdf"
HTML_FRAGMENT = "message_fragment.html"
EXCEL_FILE    = "C:\\Users\\Mike\\Dropbox\\Current Work\\Email News Contact List Addresses.xlsx"
TEST_EMAIL    = ENV['TEST_EMAIL'] || 'mjblyth@proton.me' # Configurable test email

Mail.defaults { delivery_method :smtp, SMTP_CONFIG }

# --- FUNCTIONS ---

def read_recipients(file_path)
  recipients = []
  begin
    creek = Creek::Book.new(file_path)
    sheet = creek.sheets[0] # First sheet

    sheet.simple_rows.each_with_index do |row, index|
      next if index == 0 || row.empty?
      recipients << {
        email:      row["B"],
        first_name: row["C"],
        last_name:  row["D"]
      }
    end
  rescue => e
    puts "Failed to read Excel: #{e.message}"
    exit
  end
  recipients
end

def build_html_body(first_name, fragment_path, photo_cid, doc_link)
  inner_html = File.read(fragment_path)
  <<~HTML
    <html>
      <body style="font-family: sans-serif; line-height: 1.5; color: #333; max-width: 600px;">
        <p>Hi #{first_name},</p>
        <div style="text-align: center; margin-bottom: 20px;">
          <img src="cid:#{photo_cid}" style="max-width: 100%; border-radius: 8px;">
        </div>
        #{inner_html}
        <div style="margin-top: 30px; padding: 15px; background-color: #f9f9f9; border-left: 4px solid #6d4aff;">
          <strong>Read the full story here:</strong><br>
          <a href="#{doc_link}">Open Google Doc (Preview Mode)</a>
        </div>
      </body>
    </html>
  HTML
end

def send_email(person, subject, from_email)
  begin
    Mail.deliver do
      to      person[:email]
      from    from_email
      subject subject
      add_file PHOTO_FILE
      inline_cid = attachments[PHOTO_FILE].cid
      add_file PDF_FILE
      html_part do
        content_type 'text/html; charset=UTF-8'
        body build_html_body(person[:first_name], HTML_FRAGMENT, inline_cid, DOC_LINK)
      end
    end
    puts "Email sent to #{person[:first_name]} #{person[:last_name]}."
    return true
  rescue Net::SMTPUnknownError, Net::OpenTimeout, Errno::ECONNREFUSED => e
    puts "Failed to send email to #{person[:email]}: #{e.message}"
    puts "SUGGESTION: Is Proton Bridge running and accessible at #{SMTP_CONFIG[:address]}:#{SMTP_CONFIG[:port]}?"
    return false
  rescue => e
    puts "Failed to send email to #{person[:email]}: #{e.message}"
    return false
  end
end


# --- MENU SYSTEM ---
puts "--- Newsletter Manager ---"
puts "1. Generate local HTML file for review"
puts "2. Send test email to #{TEST_EMAIL}"
puts "3. Send to WHOLE list"
print "Choose an option (1-3): "
choice = gets.chomp

case choice
when "1"
  # Generate a local file (Note: CID images won't show in local browsers)
  # To make the preview work, we'll just use the photo file path directly.
  preview_html = build_html_body("TestName", HTML_FRAGMENT, PHOTO_FILE, DOC_LINK)
  File.write("preview_full.html", preview_html)
  puts "Done! Open 'preview_full.html' in your browser."

when "2"
  puts "Sending test to #{TEST_EMAIL}..."
  test_person = { email: TEST_EMAIL, first_name: "Mike" }
  send_email(test_person, "TEST: Family Update", "mjblyth@proton.me")

when "3"
  recipients = read_recipients(EXCEL_FILE)

  puts "\nWARNING: You are about to send to #{recipients.length} people."
  print "Are you absolutely sure? (type 'YES' to broadcast): "
  confirm = gets.chomp

  if confirm == "YES"
    recipients.each do |person|
      next unless person[:email]
      puts "Broadcasting to #{person[:first_name]} #{person[:last_name]}..."
      send_email(person, "Blyth Family Christmas Letter 2025", "mjblyth@proton.me")
      sleep 2
    end
    puts "Broadcast complete."
  else
    puts "Broadcast cancelled."
  end

else
  puts "Invalid option."
end
