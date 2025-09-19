import extract_msg
import re
import csv
import os

# Path to the .msg file
msg_file_path = r"c:\Users\kreimes\Documents\Memberships and License\LREC\Secretary_Treasurer Files\Meeting message.msg"
output_file = r"c:\Users\kreimes\Documents\Memberships and License\LREC\Secretary_Treasurer Files\mailing_list.csv"

# Read the .msg file
msg = extract_msg.Message(msg_file_path)

# Get the email content
email_content = msg.body
# Also include the sender and recipients
email_content += f"\n{msg.sender}\n{msg.to}\n{msg.cc}"

# Regular expression pattern for email addresses
email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

# Find all email addresses
email_addresses = set(re.findall(email_pattern, email_content))

# Write unique email addresses to CSV
with open(output_file, 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['Email Address'])  # Header
    for email in sorted(email_addresses):
        writer.writerow([email])

print(f"Found {len(email_addresses)} unique email addresses.")
print(f"Email addresses have been saved to {output_file}")
