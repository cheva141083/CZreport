

# Use a pipeline as a high-level helper
from transformers import pipeline

pipe = pipeline("text-classification", model="BAAI/bge-reranker-v2-m3")


# Load model directly
from transformers import AutoTokenizer, AutoModelForSequenceClassification

tokenizer = AutoTokenizer.from_pretrained("BAAI/bge-reranker-v2-m3")
model = AutoModelForSequenceClassification.from_pretrained("BAAI/bge-reranker-v2-m3")

text = ("La vida es bella","la vida es fea")
resultados = pipe (text)

print(resultados)