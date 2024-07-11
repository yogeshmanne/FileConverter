import pandas as pd

# Approach 4: Using lxml parser with read_html()
def read_html_with_lxml(file_path):
	# Read HTML file into DataFrame using read_html() with 'lxml' parser
	df = pd.read_html(file_path, flavor='lxml')[0]
	return df

# File path
html_file_path = 'file:///C:/Users/Yogesh%20M/Documents/GitHub/FileConverter/fh1.html'

# Read HTML file using lxml parser with read_html()
df = read_html_with_lxml(html_file_path)

# Display DataFrame
print("Approach 4 Output:")
print(df)
