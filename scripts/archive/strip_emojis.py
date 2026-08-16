import re, json

def remove_emoji(text):
    if not isinstance(text, str):
        return text
    # Unicode ranges covering emoji symbols
    emoji_pattern = re.compile(
        "["
        "\U00010000-\U0010ffff"
        "\u2600-\u27bf"
        "\u2300-\u23ff"
        "\u2b50"
        "]+", flags=re.UNICODE
    )
    return emoji_pattern.sub(r"", text).strip()

with open("evaluation_data.json", "r", encoding="utf-8") as f:
    data = json.load(f)

def clean_obj(obj):
    if isinstance(obj, str):
        return remove_emoji(obj)
    elif isinstance(obj, list):
        return [clean_obj(x) for x in obj]
    elif isinstance(obj, dict):
        return {k: clean_obj(v) for k, v in obj.items()}
    return obj

cleaned_data = clean_obj(data)
with open("evaluation_data.json", "w", encoding="utf-8") as f:
    json.dump(cleaned_data, f, ensure_ascii=False, indent=2)

print("evaluation_data.json cleaned of all emojis!")
