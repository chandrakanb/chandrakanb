from bs4 import BeautifulSoup

# Read the HTML file
with open("MainDetailedReport.html", "r", encoding="utf-8") as file:
    content = file.read()

# Parse the HTML content
soup = BeautifulSoup(content, "html.parser")

# Extract the <head> part
head = soup.head

# Extract the specific <body> part
summary_table_div = soup.find("div", id="div-summaryTable")
pie_chart_div = soup.find("div", id="pieChartDiv")

# Create a new BeautifulSoup object for the extracted content
extracted_content = BeautifulSoup("<html></html>", "html.parser")
extracted_html = extracted_content.html

# Append the extracted head to the new BeautifulSoup object
if head:
    extracted_html.append(head.extract())

# Create and append the body with the extracted parts
body = extracted_content.new_tag("body")
extracted_html.append(body)

# Create a container div to hold the summary table and pie chart side by side
container_div = extracted_content.new_tag("div", **{"class": "container-fluid mt-5"})

# Create a row div
row_div = extracted_content.new_tag("div", **{"class": "row testPackData"})

# Create the first column for the summary table
summary_col_div = extracted_content.new_tag("div", **{"class": "col-sm-6"})
if summary_table_div:
    summary_col_div.append(summary_table_div.extract())

# Create the second column for the pie chart
chart_col_div = extracted_content.new_tag("div", **{"class": "col-sm-6"})
if pie_chart_div:
    chart_col_div.append(pie_chart_div.extract())

# Append the columns to the row
row_div.append(summary_col_div)
row_div.append(chart_col_div)

# Append the row to the container
container_div.append(row_div)

# Append the container to the body
body.append(container_div)

# Write the extracted content to a new HTML file
with open("extracted_content.html", "w", encoding="utf-8") as file:
    file.write(str(extracted_content))

print("Extraction complete. The extracted content is saved in 'extracted_content.html'.")
