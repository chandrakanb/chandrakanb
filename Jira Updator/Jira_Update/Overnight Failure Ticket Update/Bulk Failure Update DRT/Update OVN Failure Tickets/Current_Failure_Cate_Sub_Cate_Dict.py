failure_category_dict = {}

failure_category_dict = {
    "KITE issue": 16506,
    "TS issue": 16507,
    "ICB issue": 16512,
    "Test Env issue": 16508,
    "Human error": 16509
}

failure_sub_category_dict = {}

failure_category = "Human error"

if failure_category == "KITE issue":
    failure_sub_category_dict = {
        "CAN signal issue": 16522,
        "Image comparison": 16523,
        "Audio comparison": 16524,
        "Video comparison": 16525,
        "Robot operation": 16526,
        "Device Mapping issue": 16527,
        "KITE Application issue": 16528
    }
elif failure_category == "TS issue":
    failure_sub_category_dict = {
		"TS logic incorrect" : 16529
    }
elif failure_category == "ICB issue":
    failure_sub_category_dict = {
		"Once seen" : 16518,
		"Always" : 16520,
		"GAS implementation" : 16519,
		"Timing issue" : 16521,
		"SRL/CR changes" : 16554
	}
elif failure_category == "Test Env issue":
    failure_sub_category_dict = {
		"Automation Setup malfunction" : 16514,
		"External HW" : 16515,
		"External Application issue" : 16516,
		"Network Issue" : 16517,
		"Cascading Issue" : 19025
	}

print(failure_sub_category_dict)
