import json
import glob
import pandas as pd
from datetime import datetime

rows = []

json_files = glob.glob("reports/*.json")

for file in json_files:

    with open(file) as f:

        data = json.load(f)

        for item in data:

            rows.append({

                "Subscription":
                file.split("/")[-1].replace(".json", ""),

                "Category":
                item.get('category'),

                "Impact":
                item.get('impact'),

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

print("Multi Subscription Excel Report Generated")
