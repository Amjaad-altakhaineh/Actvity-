import pandas as pd
from collections import defaultdict
from nltk.stem.isri import ISRIStemmer
from nltk.corpus import stopwords
import nltk
import re

# تنزيل الموارد المطلوبة
nltk.download('stopwords')

# مسار ملف الإدخال
file_path = 'الانشطة.xlsx'

# قراءة البيانات
df = pd.read_excel(file_path, engine='openpyxl')

# حفظ النصوص الأصلية في عمود جديد
df['ORIGINAL_ACTIVITY_NAME_AR'] = df['ACTIVITY_NAME_AR']

# قاموس لتخزين الجذور والكلمات الأصلية
root_to_original = defaultdict(list)

# تهيئة ISRI Stemmer
stemmer = ISRIStemmer()

# كلمات التوقف للغة العربية
stop_words = set(stopwords.words("arabic"))
stop_words.add("و")  # إضافة "و" إلى قائمة كلمات التوقف
pattern_10= r'(بالتجزئة|بالجملة)'
pattern_11= r'(لوازمها|اللوازم)'

# معالجة النصوص وإنشاء الجذور
processed_texts = []
for text in df['ACTIVITY_NAME_AR']:

    text=text = re.sub(pattern_10, "", text)
    text=text = re.sub(pattern_11, "لوازم", text)     # إزالة الرموز غير الأبجدية
    text = re.sub(r"[^\w\s]", "", text)
    
    # تحويل أشكال الألف المختلفة إلى "ا"
    text = re.sub(r"[إأآ]", "ا", text)
    
    # إزالة التشكيل
    #text = re.sub(r"[ًٌٍَُِّْ]", "", text)
    
    # إزالة الحروف المكررة
   # text = re.sub(r"(.)\1+", r"\1", text)
    
    # إزالة حرف "و" الملاصق في بداية الكلمات
    text = re.sub(r"\bو(?=\w)", "", text)
    
    # إزالة المسافات الزائدة
    text = re.sub(r'\s+', ' ', text).strip()
    
    # إزالة كلمات التوقف
    filtered_words = [word for word in text.split() if word not in stop_words]
    
    # إنشاء خريطة الجذر إلى الكلمات الأصلية
    for word in filtered_words:
        root = stemmer.stem(word)
        root_to_original[root].append(word)
    
    # تخزين النص المعالج
    processed_texts.append(" ".join(filtered_words))

# إضافة النصوص المعالجة إلى البيانات
df['ACTIVITY_NAME_AR'] = processed_texts

# إنشاء جدول الجذور والكلمات الأصلية
mapping_data = [{"Root": root, "Original Words": ", ".join(words)} for root, words in root_to_original.items()]
mapping_df = pd.DataFrame(mapping_data)

from collections import Counter



word_list = []
for activity in df["ACTIVITY_NAME_AR"]:
    words = activity.split()

    word_list.extend( words)
    

word_counts = Counter(word_list)


top_10 = word_counts.most_common(50)
