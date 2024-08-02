import json
import pandas as pd

# Sample complex JSON structure
data = {
    "users": [
        {
            "user_id": 1,
            "name": "John Doe",
            "accounts": [
                {
                    "account_id": 101,
                    "plan": "Unlimited",
                    "devices": [
                        {"device_id": 1001, "device_name": "iPhone 12"},
                        {"device_id": 1002, "device_name": "Galaxy S21"}
                    ]
                },
                {
                    "account_id": 102,
                    "plan": "Basic",
                    "devices": [
                        {"device_id": 1003, "device_name": "Pixel 5"}
                    ]
                }
            ]
        },
        {
            "user_id": 2,
            "name": "Jane Smith",
            "accounts": [
                {
                    "account_id": 103,
                    "plan": "Family",
                    "devices": [
                        {"device_id": 1004, "device_name": "iPhone 11"}
                    ]
                }
            ]
        }
    ]
}

# Flatten the JSON while maintaining hierarchy
def flatten_json(y, parent_key='', sep='_'):
    items = []
    
    def recursive_flatten(y, parent_key=''):
        if isinstance(y, dict):
            for k, v in y.items():
                new_key = parent_key + sep + k if parent_key else k
                if isinstance(v, dict):
                    recursive_flatten(v, new_key)
                elif isinstance(v, list):
                    for i, item in enumerate(v):
                        recursive_flatten(item, new_key + sep + str(i))
                else:
                    items.append((new_key, v))
        return dict(items)
    
    return recursive_flatten(y)

# Prepare a list of all flattened entries
flattened_data = []
for user in data['users']:
    user_flat = flatten_json(user)
    flattened_data.append(user_flat)

# Convert the list of dicts to a DataFrame
df = pd.DataFrame(flattened_data)

# Save to a single Excel sheet
df.to_excel("hierarchical_data.xlsx", index=False)

print("Excel file created: hierarchical_data.xlsx")
