import requests
from bs4 import BeautifulSoup
import pandas as pd
import datetime

now = datetime.datetime.now()

url = 'https://www.cisco.com/c/en/us/support/web/tsd-products-field-notice-summary.html#~tab-most-recent'
url_jp = "https://www.cisco.com/c/ja_jp/support/web/tsd-products-field-notice-summary.html"
html = requests.get(url)
soup = BeautifulSoup(html.content, 'html.parser')

def main():
    # Convert to English month name
    month_names = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    month_name = month_names[now.month - 1]
    # Convert Date
    formatted_date = now.strftime(f'{month_name} %d, %Y')
    # print(formatted_date)
    n = 0
    data = []
    for element in soup.find_all('a', href=True):
        link = element['href']
        title = element.find('span', class_='most_recent_link_title').string if element.find('span', class_='most_recent_link_title') else None
        head = "https://www.cisco.com/"
        if title is None:
            pass
        else:
            #### Print number and description, URL #####
            n += 1
            # print(f'{n}\n{title}\n{date}\n{head}{link["href"]}')
            category = title.split(":Field Notice")[0]
            title = title.split(":Field Notice")[1]
            title = title[2:]
            # Get a Field Notice Link
            url_FN = head + link#["href"]
            html_FN = requests.get(url_FN)
            soup_FN = BeautifulSoup(html_FN.content, 'html.parser')
            # Get Date
            date = soup_FN.find('div', class_='updatedDate').text.split(':')[1].strip()
            date_history = soup_FN.find('h3', text='Revision History')
            table_elements = date_history.find_next('table')
            # print(table_element)
            # Get a date_history value
            date_elements = table_elements.find_all('div', align='center')
            # print(date_elements)
            # Get an update date
            for date_element in date_elements:
                if len(date_elements) > 2:
                    date_value = date_element.text.strip()
                    # print(f"Date Updated: {date_value}")
            # Get "Problem Description"
            problem_description_section = soup_FN.find('h3', text='Problem Description')
            #### Display text #####
            problem_description_text = problem_description_section.find_next('p').text.strip()
            # print('Problem Description:', problem_description_text)
            # Get "Problem Description" section
            workaround_section = soup_FN.find('h3', text='Workaround/Solution')
            # Workaroundのpタグを全表示してテキスト #####
            workaround_text = workaround_section.find_next_siblings('p')
            workaround_draft = ""
            for p_tag in workaround_text:
                workaround_draft += p_tag.get_text().strip() + "\n"

            # Get "Defect Information" section
            defect_section = soup_FN.find('h3', text='Defect Information')
            defect_description = defect_section.find_next('table')
            for row in defect_description.find_all('tr')[1:]:
                defect_link = row.find('a')['href']
            # print(defect_description)
            # "Products Affected"のセクションを検出
            affected_section = soup_FN.find('h3', text='Products Affected')
            affected_contents = affected_section.find_all_next('table')
            # affected_contents2 = affected_section.find_next_sibling('table')
            # print(affected_contents)

            # create a  list for table
            table_content = []
            for table in affected_contents:
                cells = table.find_all('td')
                # Get each td tag and import beside last string
                row_content = [cell.text.strip() for cell in cells]
                # Add table row to table_content
                table_content.append(row_content)

            if not table_content[0]:
                # table_content = "Software Issue: For the problem details, please check the provided link."
                table_content = '\n'.join(table_content[1])
            else:
                table_content = '\n'.join(table_content[0])

            data.append({'Num': n,
                         'Title': title,
                         'Category': category,
                         'Update Date': date,
                         'Products Affected': table_content,
                         'URL': url_FN,
                         'Description': problem_description_text,
                         'Defect URL': defect_link,
                         'Defect ID': defect_link[-10:],
                         'Workaround': workaround_draft,
                         })
            df = pd.DataFrame.from_dict(data)
            # print(df)

    # 特定Update日付だけ抽出
    df_date_output = df[
        (df["Update Date"].astype(str).str.contains(month_name))
        ]
    # print(df_date_output)
    # DataFrameをExcelファイルに出力
    writer = pd.ExcelWriter('cisco_fn' + now.strftime('%Y%m%d_%H%M%S') + '.xlsx', engine='xlsxwriter')
    # df_date_output.to_excel(writer, index=False, sheet_name='Sheet1')
    df.to_excel(writer, index=False, sheet_name='FieldNotice')
    # Excelのワークシートを取得
    workbook = writer.book
    worksheet = writer.sheets['FieldNotice']

    # Cell style
    cell_format = workbook.add_format({'align': 'left',
                                       'valign': 'vcenter',
                                       'font_name': 'Cambria',
                                       'text_wrap': True
                                       })
    cell_format_nwrap = workbook.add_format({'align': 'left',
                                       'valign': 'vcenter',
                                       'font_name': 'Cambria'
                                       })
    # Excel format
    worksheet.set_zoom(80)
    # Table range
    worksheet.add_table(0, 0, df.shape[0], df.shape[1] - 1, {'columns': [{'header': col} for col in df.columns],
                                                             'style': 'Table Style Medium 9',
                                                             'name': 'MyTable'})
    worksheet.set_column('A:A', 5, cell_format)
    worksheet.set_column('B:B', 60, cell_format)
    worksheet.set_column('C:E', 15, cell_format)
    worksheet.set_column('F:F', 15, cell_format)
    worksheet.set_column('G:H', 60, cell_format)
    worksheet.set_column('I:I', 15, cell_format)
    worksheet.set_column('J:J', 150, cell_format)
    # Excel Close
    writer.close()

if __name__ == '__main__':
    main()