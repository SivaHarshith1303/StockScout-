import requests 
import bs4
import pandas as pd 

company_name = input('Enter the Ticker name of the company : ')
url = f'https://www.screener.in/company/{company_name}/consolidated/'
result = requests.get(url)
soup = bs4.BeautifulSoup(result.text,"lxml")

thead_size = len(soup.find_all('thead'))
tbody_size = len(soup.find_all('tbody'))

lists = ['Quarterly Results','Profit & Loss','Balance Sheet','Cash Flows','Ratios','Shareholding Pattern']

# Create a Pandas Excel writer using XlsxWriter as the engine
excel_filename = f"{company_name}_data.xlsx"
with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
    for i, section_name in enumerate(lists):
        mainlist = []
        print('\n', section_name)

        # Header Part. 
        header_tr_list = soup.find_all('thead')[i].find_all('th')
        headers_columns = []
        for j in header_tr_list:
            headers_columns.append(j.text)

        # Body Part. 
        body_tr_list = soup.find_all('tbody')[i].find_all('tr')
        for j in range(len(body_tr_list)):
            body_td_list = body_tr_list[j].find_all('td')
            body_columns = []
            for k in body_td_list:
                body_columns.append(k.text)
            body_columns[0] = body_columns[0].strip()
            mainlist.append(body_columns)

        df = pd.DataFrame(mainlist, columns=headers_columns)

        # Write DataFrame to Excel sheet
        df.to_excel(writer, sheet_name=section_name, index=False)

        print(f"Data for '{section_name}' has been successfully saved to '{excel_filename}'.")

print(f"All data has been successfully saved to '{excel_filename}'.")
