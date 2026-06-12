issue_types = ['Human Error', 'ICB Issue (Always)', 'Test Env Issue', 'KITE issue', 'ICB Issue (Once)', 'TS Issue', 'SRL/CR Changes', 'KITE Issue', 'Human error']

print(issue_types)

normalized_values = []

for issue_type in issue_types:
    print(issue_type)
    normalization_map = {
        "ts issue": "TS Issue",
        "kite issue": "KITE Issue",
        "icb issue (once)": "ICB Issue (Once)",
        "icb issue (always)": "ICB Issue (Always)",
        "human error": "Human Error",
        "test env issue": "Test Env Issue",
        "srl/cr changes": "SRL/CR Changes"
    }
    
    value_lower = issue_type.lower()
    
    if value_lower in normalization_map:
        normalized_value = normalization_map[value_lower]
        if normalized_value not in normalized_values:
            normalized_values.append(normalized_value)
            #print(f"Original: {issue_type}, Normalized: {normalized_value}")
            
print(normalized_values)
	
	
	
	
	