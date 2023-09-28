import pandas as pd
from bs4 import BeautifulSoup
from fuzzywuzzy import fuzz, process
import requests

# Function to process share text
def custom_share_processing(share_text):
    return share_text.split('/', 1)[-1].strip()

# Function to convert weight string to float
def process_weight(weight_str):
    weight_str = weight_str.replace(',', '.')
    return float(weight_str) if weight_str else None

# Function to scrape fund data from a given URL
def scrape_fund_data(url, table_identifier, share_index, weight_index, fund_name, share_processing_function=None):
    shares = []
    weights = []
    
    response = requests.get(url)
    if response.status_code == 200:
        print("Successfully retrieved the webpage for", fund_name)
        soup = BeautifulSoup(response.text, 'html.parser')
        #find the table of the fund infromation in the website
        table = soup.find("table", **table_identifier)
        if table:
            rows = table.find("tbody").find_all("tr")
            for row in rows:
                columns = row.find_all("td")
                #shareindex and weight index define in which column the name of the share and the weight of the share is located
                if len(columns) >= max(share_index, weight_index):
                    share_text = columns[share_index].text.strip()
                    share = share_processing_function(share_text) if share_processing_function else share_text
                    weight = process_weight(columns[weight_index].text.strip())
                    if weight is not None:
                        shares.append(share)
                        weights.append(weight)
    
    return fund_name, pd.DataFrame({"Name": fund_name, "Share": shares, "Weight": weights})

# Function to normalize strings
def normalize_string(s):
    return ''.join(e for e in s.lower() if e.isalnum())

# Function to find the best match from a list of choices
def get_best_match(target, choices):
    best_match, score = process.extractOne(target, choices)
    return best_match if score > 80 else None

# Main function to execute the entire process
def main(fund_parameters):
    funds = {}
    for params in fund_parameters:
        fund_name, df = scrape_fund_data(**params)
        #create funds based on the scrapping data
        funds[fund_name] = df
   
    #compandf is the data upright provided in the assignment
    company_df = pd.read_excel("Netimpact.xlsx", sheet_name='500 largest companies', header=1)
    results = pd.DataFrame(columns=['Fund Name', 'Net Impact Ratio', 'Data available'])
    
    for fund_name, fund_df in funds.items():
    #    shares_and_weights_df = pd.DataFrame({
    #     'fund_name': fund_name,
    #     'share': fund_df['Share'],
    #     'weight': fund_df['Weight']
    # })
    # Print the DataFrame
        # print(shares_and_weights_df)
        total_impact = 0
        total_weight = 0  
        data_available = []
        for index, row in fund_df.iterrows():
            share_name = row['Share']
            weight = row['Weight'] / 100
            normalized_share_name = normalize_string(share_name)
            if normalized_share_name:  # Check if the normalized share name is not empty
                normalized_company_names = company_df['Company'].apply(normalize_string).tolist()
                best_match = get_best_match(normalized_share_name, normalized_company_names)
            if best_match:
                company_row = company_df[company_df['Company'].apply(normalize_string) == best_match]
                if not company_row.empty:
                    data_available.append(company_row.iloc[0]['Company'])
                    net_impact_ratio = company_row.iloc[0]['Net impact ratio'] / 100
                    total_impact += weight * net_impact_ratio
                    total_weight += weight 

        if total_weight > 0:
            print(total_weight)
            normalized_total_impact = total_impact / total_weight
        else:
            normalized_total_impact = None  

        #data available lists all the companies of the found that were found from the netimpact excel
        new_row = pd.DataFrame({'Fund Name': [fund_name], 'Net Impact Ratio': [normalized_total_impact], 'Share percentage available': [total_weight], 'Data available': [data_available]})
        results = pd.concat([results, new_row], ignore_index=True)
    
    #finally store the info 
    with pd.ExcelWriter('fund_netimpact.xlsx', engine='xlsxwriter') as writer:
        for index, row in results.iterrows():
            row_df = pd.DataFrame({'Fund Name': [row['Fund Name']], 'Net Impact Ratio': [row['Net Impact Ratio']], 'Share percentage available' : [row['Share percentage available']], 'Data available': [row['Data available']]})
            sheet_name = row['Fund Name'][:30]  # Truncate sheet name to avoid Excel limitations
            row_df.to_excel(writer, sheet_name=sheet_name, index=False)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            percent_format = workbook.add_format({'num_format': '0.00%'})
            worksheet.set_column('B:B', None, percent_format)
            worksheet.set_column('C:C', None, percent_format)
            
    # print results to get a quick overview
    print(results)

# Define fund parameters
fund_parameters = [
    {
        "url": "https://www.seligson.fi/luotain/FundViewer.php?view=10&lang=0&fid=795&task=intro",
        "table_identifier": {"class": "fundprobe company"},
        "share_index": 0,
        "weight_index": 4,
        "fund_name": "Seligson & Co Global Top 25 Brands"
    },
    {
        "url": "https://fintel.io/i/ishares-trust-ishares-core-s-p-500-etf",
        "table_identifier": {"id": "transactions"},
        "share_index": 0,
        "weight_index": 8,
        "fund_name": "iShares Core S&P 500 ETF",
        "share_processing_function": custom_share_processing
    },
    {
        "url": "https://fintel.io/i/flexshares-trust-flexshares-stoxx-global-esg-impact-index-fund",
        "table_identifier": {"id": "transactions"},
        "share_index": 0,
        "weight_index": 8,
        "fund_name": "STOXX Global ESG Impact",
        "share_processing_function": custom_share_processing
    }
]

# Execute main function
if __name__ == "__main__":
    main(fund_parameters)
