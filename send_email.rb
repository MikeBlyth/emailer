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
  openssl_verify_mode:  'none'
}

DOC_LINK      = "https://bit.ly/3YaNO35"
PHOTO_FILE    = "header_photo.jpg"
PDF_FILE      = "Christmas Letter 2025.pdf"
HTML_FRAGMENT = "message_fragment.html"
EXCEL_FILE    = "C:\\Users\\Mike\\Dropbox\\Current Work\\Email News Contact List Addresses.xlsx"

Mail.defaults { delivery_method :smtp, SMTP_CONFIG }

# --- READ EXCEL ---
begin
  creek = Creek::Book.new(EXCEL_FILE)
  sheet = creek.sheets[0] # First sheet
  
  # Map rows to a cleaner format
  # Column 2 (B) = Email, 3 (C) = First Name, 4 (D) = Last Name
  recipients = []
  
  # We skip the first row (headers) by using .each_with_index or a simple counter
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

# --- MENU SYSTEM ---
puts "--- Newsletter Manager ---"
puts "1. Generate local HTML file for review"
puts "2. Send test email to YOURSELF only"
puts "3. Send to WHOLE list"
print "Choose an option (1-3): "
choice = gets.chomp

case choice
when "1"
  # Generate a local file (Note: CID images won't show in local browsers)
  full_html = build_html_body("TestName", HTML_FRAGMENT, "local_photo", DOC_LINK)
  File.write("preview_full.html", full_html)
  puts "Done! Open 'preview_full.html' in your browser. (Note: The inline photo will look broken in a local browser; it only works in email clients)."

when "2"
  puts "Sending test to mjblyth@proton.me..."
  Mail.deliver do
    to      'mjblyth@proton.me'
    from    'mjblyth@proton.me'
    subject "TEST: Family Update"
    add_file PHOTO_FILE
    inline_cid = attachments[PHOTO_FILE].cid
    add_file PDF_FILE
    html_part do
      content_type 'text/html; charset=UTF-8'
      body build_html_body("Mike", HTML_FRAGMENT, inline_cid, DOC_LINK)
    end
  end
  puts "Test email sent."

when "3"
  # READ EXCEL
  creek = Creek::Book.new(EXCEL_FILE)
  sheet = creek.sheets[0]
  recipients = []
  sheet.simple_rows.each_with_index do |row, index|
    next if index == 0 || row.empty?
    recipients << { email: row["B"], first_name: row["C"], last_name: row["D"] }
  end

  puts "\nWARNING: You are about to send to #{recipients.length} people."
  print "Are you absolutely sure? (type 'YES' to broadcast): "
  confirm = gets.chomp

  if confirm == "YES"
    recipients.each do |person|
      next unless person[:email]
      puts "Broadcasting to #{person[:first_name]} #{person[:last_name]}..."
      Mail.deliver do
        to      person[:email]
        from    'mjblyth@proton.me'
        subject "Blyth Family Christmas Letter 2025"
        add_file PHOTO_FILE
        inline_cid = attachments[PHOTO_FILE].cid
        add_file PDF_FILE
        html_part do
          content_type 'text/html; charset=UTF-8'
          body build_html_body(person[:first_name], HTML_FRAGMENT, inline_cid, DOC_LINK)
        end
      end
      sleep 2
    end
    puts "Broadcast complete."
  else
    puts "Broadcast cancelled."
  end

else
  puts "Invalid option."
end
