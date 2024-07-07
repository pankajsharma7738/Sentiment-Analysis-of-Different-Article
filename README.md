Instructions for Running the Python Script
Ensure you have the following files in the same directory:
● text_analysis.py
● Input.xlsx
● stopwords directory (containing stopword text files)
● master directory (containing positive-words.txt and negative-words.txt)
Run the Script:
● Open a terminal or command prompt.
● Navigate to the directory containing your script and files.
● Execute the script using Python : python text_analysis.py
Output:
● The script will process the articles and generate two output files:
○ Cleaned article text files in the articles directory.
○ An Excel file named Output.csv containing the calculated variables.
Dependencies
Install required Python packages :
● RUN THE CODE : pip install pandas requests beautifulsoup4 nltk openpyxl
pandas: For data manipulation and analysis.
requests: For making HTTP requests to fetch web pages.
beautifulsoup4: For parsing HTML content.
nltk: For natural language processing tasks.
openpyxl: For reading and writing Excel files.
