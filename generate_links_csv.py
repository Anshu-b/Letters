import csv

# Input CSV file
input_csv = 'judges.csv'

# Output CSV file
output_csv = 'judge_links.csv'

# Read the input CSV and write to output CSV
with open(input_csv, 'r') as infile:
    # Skip the first line
    infile.readline()
    reader = csv.DictReader(infile)
    
    with open(output_csv, 'w', newline='') as outfile:
        fieldnames = ['Track', 'Name', 'Email', 'Certificate Link', 'Letter Link']
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        
        for row in reader:
            name = row['Name'].strip()
            track = row['Track'].strip()
            email = row['Email'].strip()
            
            # Generate links
            name_underscored = name.replace(' ', '_')
            cert_link = f'https://github.com/Anshu-b/Letters/blob/main/certificates/{name_underscored}.pdf'
            letter_link = f'https://github.com/Anshu-b/Letters/blob/main/letters/{name_underscored}_letter.docx'
            
            writer.writerow({
                'Track': track,
                'Name': name,
                'Email': email,
                'Certificate Link': cert_link,
                'Letter Link': letter_link
            })

print(f"Generated {output_csv} with judge links.")