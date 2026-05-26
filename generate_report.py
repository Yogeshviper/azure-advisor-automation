import json
import pandas as pd
from datetime import datetime

with open('advisor.json') as f:
    data = json.load(f)

rows = []

for item in data:

    rows.append({

        "Category": item.get('category'),

        "Impact": item.get('impact'),

        "Recommendation":
        item.get('shortDescription', {}).get('problem'),

        "Resource":
        item.get('resourceMetadata', {}).get('resourceId')

    })

df = pd.DataFrame(rows)

filename = f"Azure_Advisor_Report_{datetime.now().date()}.xlsx"

with pd.ExcelWriter(filename, engine='openpyxl') as writer:

    df.to_excel(
        writer,
        sheet_name='Advisor Report',
        index=False
    )

print("Excel Report Generated")