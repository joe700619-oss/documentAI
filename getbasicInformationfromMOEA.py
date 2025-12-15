import requests
import json

class BasicInformationAPI:
    def __init__(self):
        # API 1: Company Basic Information
        self.basic_info_url = "https://data.gcis.nat.gov.tw/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6"
        # API 2: Company Business Items
        self.business_items_url = "https://data.gcis.nat.gov.tw/od/data/api/236EE382-4942-41A9-BD03-CA0709025E7C"
        # API 3: Company Directors
        self.directors_url = "https://data.gcis.nat.gov.tw/od/data/api/4E5F7653-1B91-4DDC-99D5-468530FAE396"

    def _fetch_from_api(self, url, business_accounting_no):
        params = {
            "$format": "json",
            "$filter": f"Business_Accounting_NO eq {business_accounting_no}"
        }
        try:
            response = requests.get(url, params=params)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            print(f"Error fetching API data from {url}: {e}")
            return []

    def get_company_facts(self, business_accounting_no):
        # 1. Basic Info
        raw_basic = self._fetch_from_api(self.basic_info_url, business_accounting_no)
        basic_data = {}
        if raw_basic and isinstance(raw_basic, list) and len(raw_basic) > 0:
            info = raw_basic[0]
            basic_data = {
                "companyName": info.get("Company_Name", ""),
                "authorizedCapital": info.get("Capital_Stock_Amount", 0),
                "companyAddress": info.get("Company_Location", ""),
                "chairmanName": info.get("Responsible_Name", "")
            }

        # 2. Business Items
        raw_items = self._fetch_from_api(self.business_items_url, business_accounting_no)
        business_items = []
        if raw_items and isinstance(raw_items, list) and len(raw_items) > 0:
            items_info = raw_items[0]
            cmp_business_list = items_info.get("Cmp_Business", [])
            for item in cmp_business_list:
                # Provide id, code, and name as per data.json style
                seq = item.get("Business_Seq_NO", "")
                # Remove leading zeros if it's purely numeric
                if seq.isdigit():
                    seq = str(int(seq))
                
                business_items.append({
                    "id": seq,
                    "code": item.get("Business_Item", ""),
                    "name": item.get("Business_Item_Desc", "")
                })

        # 3. Directors
        # raw_directors = self._fetch_from_api(self.directors_url, business_accounting_no)
        directors = []
        # if raw_directors and isinstance(raw_directors, list):
        #     for i, item in enumerate(raw_directors, 1):
        #         shares_val = item.get("Person_Shareholding", 0)
        #         shares_str = f"{shares_val:,}股" if isinstance(shares_val, (int, float)) else f"{shares_val}股"
        #         
        #         directors.append({
        #             "id": str(i),
        #             "position": item.get("Person_Position_Name", ""),
        #             "name": item.get("Person_Name", ""),
        #             "shares": shares_str,
        #             "id_number": "",  # API does not provide this
        #             "address": ""    # API does not provide this
        #         })

        # Combine
        facts = {
            **basic_data,
            "business_items": business_items,
            "directors": directors
        }
        return facts

def main():
    api = BasicInformationAPI()
    target_no = "60299784"
    print(f"Fetching data for Business Accounting NO: {target_no}")
    
    facts = api.get_company_facts(target_no)
    
    # Generate granular policy for directors
    # Fields from API are locked, missing fields are editable
    directors_policy = []
    for _ in facts.get("directors", []):
        directors_policy.append({
            "id": "editable",
            "position": "editable",
            "name": "editable",
            "shares": "editable",
            "id_number": "editable",
            "address": "editable"
        })

    field_policy = {
        "companyName": "locked",
        "authorizedCapital": "locked",
        "companyAddress": "locked",
        "chairmanName": "locked",
        "business_items": "locked",
        "directors": directors_policy
    }

    output_data = {
        "company_facts_authoritative": facts,
        "field_policy": field_policy
    }
    
    output_file = "test_data.json"
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        print(f"Data successfully written to {output_file}")
        # Print a snippet for verification
        print(json.dumps(output_data, ensure_ascii=False, indent=2)[:500] + "...")
    except IOError as e:
        print(f"Error writing to file {output_file}: {e}")

if __name__ == "__main__":
    main()
