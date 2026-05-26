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

            extended = item.get('extendedProperties', {})

            resource_metadata = item.get('resourceMetadata', {})

            short_desc = item.get('shortDescription', {})

            rows.append({

                "Category":
                item.get('category'),

                "Business Impact":
                item.get('impact'),

                "Recommendation":
                short_desc.get('problem'),

                "Subscription ID":
                resource_metadata.get('subscriptionId'),

                "Subscription Name":
                file.split("/")[-1].replace(".json", ""),

                "Resource Group":
                resource_metadata.get('resourceGroup'),

                "Resource Name":
                resource_metadata.get('resourceName'),

                "Type":
                resource_metadata.get('resourceType'),

                "Last Refreshed":
                item.get('lastUpdated'),

                "Potential Benefits":
                short_desc.get('solution'),

                "Potential Annual Cost Savings":
                extended.get('annualSavingsAmount'),

                "Potential Cost Savings Currency":
                extended.get('savingsCurrency'),

                "Cost Implications":
                extended.get('costImplication'),

                "Description of Changes":
                extended.get('recommendedActions')
            })

df = pd.DataFrame(rows)

filename = f"Azure_Advisor_Report_{datetime.now().date()}.xlsx"

with pd.ExcelWriter(filename, engine='openpyxl') as writer:

    df.to_excel(
        writer,
        sheet_name='Advisor Report',
        index=False
    )

print("Enhanced Azure Advisor Report Generated")
