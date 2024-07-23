import pandas as pd
from atlassian import Confluence

# Step 1: Read and process the Excel file
def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        if df.empty:
            print("The Excel file is empty.")
            return None
        return df
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return None

# Convert DataFrame to HTML table
def dataframe_to_html(df):
    return df.to_html(index=False)

def main():
    # Configuration
    excel_file_path = 'path/to/your/excel_file.xlsx'
    confluence_url = 'https://learniverse9618.atlassian.net/wiki'
    username = 'your-confluence-username'
    api_token = 'your-confluence-api-token'
    page_id = 22642689  # Confluence page ID
    
    # Read and process the Excel file
    df = read_excel(excel_file_path)
    if df is None:
        print("Failed to read the Excel file. Exiting...")
        return
    
    # Convert DataFrame to HTML
    content = dataframe_to_html(df)
    
    # Authenticate with Confluence
    try:
        confluence = Confluence(
            url=confluence_url,
            username=username,
            password=api_token
        )
        
        # Verify authentication by fetching spaces
        spaces = confluence.get_all_spaces(start=0, limit=1)
        if 'statusCode' in spaces and spaces['statusCode'] == 401:
            print("Authentication failed: Check your credentials.")
            return
    except Exception as e:
        print(f"Failed to authenticate with Confluence: {e}")
        return
    
    # Fetch the current page content
    try:
        page = confluence.get_page_by_id(page_id, expand='body.storage')
        if not page:
            print("Confluence page not found.")
            return
    except Exception as e:
        print(f"Error fetching the Confluence page: {e}")
        return
    
    # Get the current page body
    page_body = page.get("body").get("storage").get("value")
    
    # Update the page body with the new content
    page_body = content
    
    # Update the Confluence page
    try:
        confluence.update_page(
            page_id=page_id,
            title=page['title'],
            body=page_body
        )
        print("Page updated successfully:", page['title'])
    except Exception as e:
        print(f"Error updating the Confluence page: {e}")

if __name__ == "__main__":
    main()
