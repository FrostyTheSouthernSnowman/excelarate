from excellent import SpreadSheet

sheet = SpreadSheet("Create table example")
sheet.add_dict({"some": "data", 1: "number keys", "number values": 4, "lists": [
               "some", "more", "data"], "nested dicts": {"still": "work"}}, 3, 5)
sheet.save_sheet("simple_table.xlsx")
