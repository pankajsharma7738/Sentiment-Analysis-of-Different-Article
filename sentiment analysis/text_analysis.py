import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
from nltk.tokenize import word_tokenize, sent_tokenize
import re

# Function to fetch article title and text
def fetch_article(url):
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        title = soup.find('title').get_text()
        
        paragraphs = []
        start_scraping = False
        for p in soup.find_all('p'):
            text = p.get_text()
            
            # Start scraping from the <p> tag containing "Client :"
            if re.search(r"Client\s*:", text):
                start_scraping = True
            
            if start_scraping:
                if not re.search(r"Summarized:|Firm Name:|Firm Website:|Firm Address:|Email:|Skype:|WhatsApp:|Telegram:|Contact us:|This project was done by|© All Right Reserved|We provide intelligence, accelerate innovation and", text):
                    paragraphs.append(text)
        
        article_text = ' '.join(paragraphs)
        return title, article_text.strip()
    else:
        return None, None

# Function to read all stop words from stop words directory
def get_stopwords(stopwords_dir):
    stopwords = set()
    for stopwords_file in os.listdir(stopwords_dir):
        if stopwords_file.endswith('.txt'):  # Ensure only text files are read
            with open(os.path.join(stopwords_dir, stopwords_file), 'r', encoding='utf-8', errors='ignore') as file:
                content = file.read().replace(' | ', '\n').replace('\n', ' ')
                words = content.split()
                stopwords.update(word.lower() for word in words)  # Convert to lowercase
    return stopwords

# Function to remove stop words from text
def remove_stopwords(text, stopwords):
    words = text.split()
    cleaned_text = ' '.join([word for word in words if word.lower() not in stopwords])
    return cleaned_text

# Function to calculate positive and negative scores using master dictionary
def calculate_positive_negative_scores(text, positive_words, negative_words):
    tokens = word_tokenize(text.lower())
    
    # Calculate positive and negative scores
    positive_score = sum(1 for word in tokens if word in positive_words)
    negative_score = sum(1 for word in tokens if word in negative_words)
    
    return positive_score, negative_score

# Function to calculate derived variables
def calculate_variables(text, positive_words, negative_words):
    tokens = word_tokenize(text)
    sentences = sent_tokenize(text)
    
    # Polarity Score
    positive_score, negative_score = calculate_positive_negative_scores(text, positive_words, negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    
    # Subjectivity Score
    subjectivity_score = (positive_score + negative_score) / (len(tokens) + 0.000001)
    
    # Readability Analysis
    avg_sentence_length = len(tokens) / len(sentences)
    complex_words = [word for word in tokens if len([char for char in word if char in 'aeiouAEIOU']) > 2]
    percentage_of_complex_words = len(complex_words) / len(tokens)
    fog_index = 0.4 * (avg_sentence_length + percentage_of_complex_words)
    avg_words_per_sentence = len(tokens) / len(sentences)
    syllable_count_per_word = sum(len([char for char in word if char in 'aeiouAEIOU']) for word in tokens) / len(tokens)
    personal_pronouns = len(re.findall(r'\b(I|we|my|ours|us)\b', text, re.I))
    avg_word_length = sum(len(word) for word in tokens) / len(tokens)
    
    return {
        'Polarity Score': polarity_score,
        'Subjectivity Score': subjectivity_score,
        'Avg Sentence Length': avg_sentence_length,
        'Percentage of Complex Words': percentage_of_complex_words,
        'Fog Index': fog_index,
        'Avg Words Per Sentence': avg_words_per_sentence,
        'Syllable Count Per Word': syllable_count_per_word,
        'Personal Pronouns': personal_pronouns,
        'Avg Word Length': avg_word_length
    }

# Main script to process articles
def main(input_file='Input.xlsx', output_file='Output.xlsx', output_dir='articles', stopwords_dir='stopwords', master_dict_dir='master'):
    # Read URLs from Excel File
    df = pd.read_excel(input_file)
    
    # Check if there are any URLs in the input file
    if len(df) == 0:
        print("No URLs found in the input file.")
        return
    
    # Create directory for articles if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Read master dictionary files and stopwords
    def read_master_dict(filename):
        with open(os.path.join(master_dict_dir, filename), 'r', encoding='utf-8', errors='ignore') as file:
            return set(file.read().split())
    
    positive_words = read_master_dict('positive-words.txt').difference(get_stopwords(stopwords_dir))
    negative_words = read_master_dict('negative-words.txt').difference(get_stopwords(stopwords_dir))
    stopwords = get_stopwords(stopwords_dir)
    
    results = []
    
    for index, row in df.iterrows():
        url_id = row['URL_ID']
        url = row['URL']
        
        # Fetch article, remove stopwords, and calculate variables
        title, article_text = fetch_article(url)
        
        if title and article_text:
            cleaned_article_text = remove_stopwords(article_text, stopwords)
            
            # Save cleaned article text to file
            article_file_path = os.path.join(output_dir, f"{url_id}_cleaned.txt")
            with open(article_file_path, 'w', encoding='utf-8') as f:
                f.write(cleaned_article_text)
            
            # Calculate positive and negative scores
            pos_score, neg_score = calculate_positive_negative_scores(cleaned_article_text, positive_words, negative_words)
            
            # Calculate derived variables
            variables = calculate_variables(cleaned_article_text, positive_words, negative_words)
            
            # Prepare row for DataFrame
            row_data = {
                'URL_ID': url_id,
                'URL': url,
                'POSITIVE SCORE': pos_score,
                'NEGATIVE SCORE': neg_score,
                'POLARITY SCORE': variables['Polarity Score'],
                'SUBJECTIVITY SCORE': variables['Subjectivity Score'],
                'AVG SENTENCE LENGTH': variables['Avg Sentence Length'],
                'PERCENTAGE OF COMPLEX WORDS': variables['Percentage of Complex Words'],
                'FOG INDEX': variables['Fog Index'],
                'AVG NUMBER OF WORDS PER SENTENCE': variables['Avg Words Per Sentence'],
                'COMPLEX WORD COUNT': len([word for word in word_tokenize(cleaned_article_text) if len([char for char in word if char in 'aeiouAEIOU']) > 2]),
                'WORD COUNT': len(word_tokenize(cleaned_article_text)),
                'SYLLABLE PER WORD': variables['Syllable Count Per Word'],
                'PERSONAL PRONOUNS': variables['Personal Pronouns'],
                'AVG WORD LENGTH': variables['Avg Word Length']
            }
            
            results.append(row_data)
            
            print(f"Processed URL_ID: {url_id} - {url}")
        else:
            print(f'Failed to fetch article for URL_ID: {url_id}')
    
    # Create DataFrame from results
    results_df = pd.DataFrame(results)
    
    # Save results to Excel file
    results_df.to_excel(output_file, index=False)
    
    print(f"Output saved to {output_file}")
    print("Processing complete.")

# Execute main script
if __name__ == "__main__":
    main()
