import re
import json
import pandas as pd

# Load file contents
with open('example/result.json') as f:
    json_content = json.dumps(json.load(f), indent=2)

with open('example/script.yaml') as f:
    yaml_content = f.read()

csv_df = pd.read_csv('example/result.csv')
csv_content = csv_df.to_markdown(index=False)

# Read README
with open('README.md', 'r') as f:
    readme = f.read()

# Replacement patterns
patterns = {
    'json': json_content,
    'yaml': yaml_content,
    'csv': csv_content
}

def replace_code_block(content, lang, new_text):
    pattern = re.compile(rf'(```{lang}\n)(.*?)(```)', re.DOTALL)
    return pattern.sub(rf'\1{new_text}\n\3', content)

for lang, new_text in patterns.items():
    readme = replace_code_block(readme, lang, new_text)

# Write updated README
with open('README.md', 'w') as f:
    f.write(readme)
