from bs4 import BeautifulSoup

# Read the content of the HTML file
with open('MainDetailedReport.html', 'r', encoding='utf-8') as file:
    html_content = file.read()

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

# Find the start time
start_time_tag = soup.find('th', string='Start Time')
start_time_value = start_time_tag.find_next_sibling('td').text if start_time_tag else None

# Split the start time into date and time components
if start_time_value:
    date, time = start_time_value.split(' ', 3)[:3], start_time_value.split(' ')[-1]
    date = ' '.join(date)
    time_hhmm = ':'.join(time.split(':')[:2])

    # Print the date and time variables
    print(f'Date: {date}')
    print(f'Time: {time_hhmm}')
else:
    print('Start Time not found')
