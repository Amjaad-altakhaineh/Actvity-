import pandas as pd
from collections import defaultdict
from nltk.stem.isri import ISRIStemmer
from nltk.corpus import stopwords
import nltk
import re

# Download required NLTK resources
nltk.download('stopwords')

# Path to the input Excel file
file_path = 'الانشطة.xlsx'

# Read the data
df = pd.read_excel(file_path, engine='openpyxl')

# Preserve original names in a new column
df['ORIGINAL_ACTIVITY_NAME_AR'] = df['ACTIVITY_NAME_AR']

# Dictionary to store the mapping between roots and original words
root_to_original = defaultdict(list)

# Initialize stemmer
stemmer = ISRIStemmer()

# Stopwords for Arabic
stop_words = set(stopwords.words("arabic"))
pattern_10= r'(بالتجزئة|بالجملة)'
# Normalize Arabic text and create root-to-original mapping
processed_texts = []
for text in df['ACTIVITY_NAME_AR']:
    text=text = re.sub(pattern_10, "", text)
    
    # Remove non-alphanumeric characters
    text = re.sub(r"[^\w\s]", "", text)
    
    # Normalize Arabic letters (e.g., convert different forms of Alef to 'ا')
    text = re.sub(r"[إأآ]", "ا", text)
    
    # Remove diacritics
    text = re.sub(r"[ًٌٍَُِّْ]", "", text)
    
    # Remove repeated characters (e.g., "رااائع" becomes "رائع")
    text = re.sub(r"(.)\1+", r"\1", text)
    
    # Remove extra spaces and strip leading/trailing whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Remove stopwords
    filtered_words = [word for word in text.split() if word not in stop_words]
    
    # Map roots to original words
    for word in filtered_words:
        root = stemmer.stem(word)
        root_to_original[root].append(word)
    
    # Store normalized text
    processed_texts.append(" ".join([stemmer.stem(word) for word in filtered_words]))

# Add normalized texts to the dataframe
df['ACTIVITY_NAME_AR'] = processed_texts

# Create the mapping DataFrame
mapping_data = [{"Root": root, "Original Words": ", ".join(words)} for root, words in root_to_original.items()]
mapping_df = pd.DataFrame(mapping_data)

# Display the mapping table to the user
import ace_tools as tools; tools.display_dataframe_to_user(name="Root to Original Words Mapping", dataframe=mapping_df)



print(mapping_df[mapping_df["Root"]=="واد"]["Original Words"])






salary_data[salary_data['EmployeeName'] == 'ALBERT PARDINI']['JobTitle']

from collections import Counter



word_list = []
for activity in df["ACTIVITY_NAME_AR"]:
    words = activity.split()

    word_list.extend( words)
    

word_counts = Counter(word_list)

top_10 = word_counts.most_common(50)
word_counts_df = pd.DataFrame(word_counts.items(), columns=['Word', 'Frequency']).sort_values(by='Frequency', ascending=False)



pattern_4= r'(أجهزة|الأجهزة|اجهزة)'


final=[]
for i in data['Joined_Text']:
    modified_text_4= re.sub(pattern_4, 'أجهزة', i, flags=re.IGNORECASE)
    final.append(modified_text_4)




new_after_clean=pd.DataFrame(final)


new_after_clean["Joined_Text"]=new_after_clean


# Step 6: Count most frequent words
word_list = []
for activity in new_after_clean["Joined_Text"]:
    words = activity.split()  # Split activity into words
    # Remove stop words and strip prefix "و"

    word_list.extend( words)

# Count word frequencies
word_counts = Counter(word_list)

word_counts_df = pd.DataFrame(word_counts.items(), columns=['Word', 'Frequency']).sort_values(by='Frequency', ascending=False)

filtered_activities =data[data['ACTIVITY_NAME_AR'].str.startswith(('تجارة', 'التجارة'), na=False)]
filtered_activities.to_excel("filtered_activities.xlsx", index=False, engine='openpyxl')
print("Filtered activities saved to 'filtered_activities.xlsx'.")
print(filtered_activities.head())

processed_activities=filtered_activities["Joined_Text"].tolist()

# Compute cosine similarity and identify unique activities
vectorizer = TfidfVectorizer()
tfidf_matrix = vectorizer.fit_transform(processed_activities)

# Calculate similarity matrix and identify duplicates
similarity_matrix = cosine_similarity(tfidf_matrix, tfidf_matrix)
threshold = 0.4
duplicates = set()
for i in range(len(similarity_matrix)):
    for j in range(i + 1, len(similarity_matrix)):
        if similarity_matrix[i, j] > threshold:
            duplicates.add((i, j))

# Extract unique activities
unique_indices = list(set(range(len(processed_activities))) - {j for _, j in duplicates})
unique_activities = [filtered_activities.iloc[i]['Joined_Text'] for i in unique_indices]

# Create results DataFrame
results = pd.DataFrame({
    "Original": filtered_activities['Joined_Text'].tolist(),
    "After": processed_activities,
    "Is unique": [i in unique_indices for i in range(len(filtered_activities))]
})


# Save unique activities to an Excel file
unique_results = pd.DataFrame({"Unique Activities": unique_activities})
unique_results.to_excel("unique_activities.xlsx", index=False, engine='openpyxl')
print("Unique activities saved to 'unique_activities.xlsx'.")
print(unique_results.head())

# Step 6: Count most frequent words
word_list = []
for activity in unique_results["Unique Activities"]:
    words = activity.split()  # Split activity into words
    # Remove stop words and strip prefix "و"

    word_list.extend( words)

# Count word frequencies
word_counts = Counter(word_list)

# Convert counts to a DataFrame and save to Excel


word_counts_df = pd.DataFrame(word_counts.items(), columns=['Word', 'Frequency']).sort_values(by='Frequency', ascending=False)
word_counts_df.to_excel("word_frequencies_3.xlsx", index=False, engine='openpyxl')



filtered_rows_mac = unique_results[
    unique_results['Unique Activities'].str.contains('معدات', na=False)
]

# Step 6: Count most frequent words
word_list = []
for activity in filtered_rows_mac["Unique Activities"]:
    words = activity.split()  # Split activity into words
    # Remove stop words and strip prefix "و"

    word_list.extend( words)

# Count word frequencies
word_counts = Counter(word_list)

word_counts_df = pd.DataFrame(word_counts.items(), columns=['Word', 'Frequency']).sort_values(by='Frequency', ascending=False)


After_cleaning_2=[]

for i in range(0,718):
    u= unique_results['Unique Activities'][i]
    o=u.split()
    if "معدات" in o :
     final_3 = ' '.join(o)
     After_cleaning_2.append(final_3)


import pandas as pd
from collections import Counter
import re
