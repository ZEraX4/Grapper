from rapidfuzz import process, fuzz

choices = ["hello world", "goodbye world", "fuzzy wuzzy"]
query = "hello"

results = process.extract(query, choices, scorer=fuzz.partial_ratio, limit=2)
print("Results:", results)
for res in results:
    print(f"Type: {type(res)}, Length: {len(res)}")
    print(res)
