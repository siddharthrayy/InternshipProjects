import re

def calculate_average_sentence_length(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()

            # Split the text into sentences using regex
            sentences = re.split(r'[.!?]', text)
            # Remove empty sentences
            sentences = [sentence.strip() for sentence in sentences if sentence.strip()]

            total_sentences = len(sentences)
            total_words = 0

            for sentence in sentences:
                words = sentence.split()
                total_words += len(words)

            average_length = total_words / total_sentences
            return average_length

    except FileNotFoundError:
        print("File not found.")
    except IOError:
        print("Error reading file.")

# Example usage:
file_path = 'scraped_data.txt'  # Replace with the path to your text file

average_length = calculate_average_sentence_length(file_path)
if average_length:
    print(f"Average sentence length: {average_length:.2f}")
