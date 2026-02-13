import os
import re
from bs4 import BeautifulSoup
import pandas as pd
from pathlib import Path

def extract_table_metadata(html_file):
    """Extract comprehensive metadata from a single HTML table documentation file"""
    
    with open(html_file, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')
    
    # Extract table name from title
    title = soup.find('title')
    table_name = title.text.replace('dbo.', '') if title else Path(html_file).stem
    
    # Extract row count
    row_count = 'N/A'
    row_count_td = soup.find('td', text=re.compile(r'Row Count', re.IGNORECASE))
    if row_count_td:
        count_cell = row_count_td.find_next_sibling('td')
        if count_cell:
            row_count = count_cell.text.strip().replace(',', '')
    
    # Extract columns section
    primary_keys = []
    columns = []
    indexed_columns = []
    
    # Find the columns table
    columns_section = soup.find('a', {'name': 'columns'})
    if columns_section:
        table = columns_section.find_next('table', {'class': 'grid'})
        if table:
            rows = table.find_all('tr')
            
            for row in rows[1:]:  # Skip header
                cells = row.find_all('td')
                if len(cells) >= 7:
                    # Check for PK and Index icons
                    icons_cell = cells[0]
                    is_pk = 'pkcluster.png' in str(icons_cell) or 'pk.png' in str(icons_cell)
                    is_indexed = 'Index.png' in str(icons_cell)
                    
                    col_name = cells[1].text.strip()
                    data_type = cells[2].text.strip()
                    max_length = cells[3].text.strip()
                    nullable = cells[4].text.strip()
                    is_identity = cells[5].text.strip()
                    default_value = cells[6].text.strip()
                    
                    columns.append({
                        'column_name': col_name,
                        'data_type': data_type,
                        'max_length': max_length,
                        'nullable': nullable,
                        'is_identity': is_identity,
                        'default_value': default_value,
                        'is_primary_key': is_pk,
                        'is_indexed': is_indexed
                    })
                    
                    if is_pk:
                        primary_keys.append(col_name)
                    if is_indexed:
                        indexed_columns.append(col_name)
    
    # Extract description/comments
    description = ''
    desc_section = soup.find('td', text=re.compile(r'Description', re.IGNORECASE))
    if desc_section:
        desc_cell = desc_section.find_next_sibling('td')
        if desc_cell:
            description = desc_cell.text.strip()
    
    # Extract indexes
    indexes = []
    index_section = soup.find('a', {'name': 'indexes'})
    if index_section:
        index_table = index_section.find_next('table', {'class': 'grid'})
        if index_table:
            idx_rows = index_table.find_all('tr')[1:]
            for idx_row in idx_rows:
                idx_cells = idx_row.find_all('td')
                if len(idx_cells) >= 2:
                    indexes.append({
                        'index_name': idx_cells[0].text.strip(),
                        'columns': idx_cells[1].text.strip() if len(idx_cells) > 1 else ''
                    })
    
    # Extract foreign keys/relationships
    foreign_keys = []
    fk_section = soup.find('a', {'name': 'foreignkeys'})
    if fk_section:
        fk_table = fk_section.find_next('table', {'class': 'grid'})
        if fk_table:
            fk_rows = fk_table.find_all('tr')[1:]
            for fk_row in fk_rows:
                fk_cells = fk_row.find_all('td')
                if len(fk_cells) >= 3:
                    foreign_keys.append({
                        'fk_name': fk_cells[0].text.strip(),
                        'referencing_column': fk_cells[1].text.strip() if len(fk_cells) > 1 else '',
                        'referenced_table': fk_cells[2].text.strip() if len(fk_cells) > 2 else ''
                    })
    
    return {
        'table_name': table_name,
        'row_count': row_count,
        'description': description,
        'primary_keys': ', '.join(primary_keys),
        'column_count': len(columns),
        'columns': columns,
        'indexed_columns': ', '.join(indexed_columns),
        'indexes': indexes,
        'foreign_keys': foreign_keys
    }

def categorize_table(table_name):
    """Categorize table based on naming convention"""
    name_upper = table_name.upper()
    
    if name_upper.startswith('SAS_'):
        if 'RATES' in name_upper:
            return 'SAS - Rates'
        elif 'LAND' in name_upper or 'LOT' in name_upper:
            return 'SAS - Land/Lot'
        elif 'PIC' in name_upper:
            return 'SAS - PIC (Property ID)'
        elif 'OWNER' in name_upper or 'CUSTOMER' in name_upper:
            return 'SAS - Customer/Owner'
        elif 'REPORT' in name_upper:
            return 'SAS - Reporting'
        elif 'WACUSER' in name_upper or 'RIGHT' in name_upper:
            return 'SAS - Security/Users'
        elif 'HIST' in name_upper:
            return 'SAS - History'
        else:
            return 'SAS - Other'
    elif 'CUSTOMER' in name_upper:
        return 'Customer Management'
    elif 'OBSERVATION' in name_upper or 'LOCUST' in name_upper or 'PEST' in name_upper or 'POISON' in name_upper:
        return 'Pest Management'
    elif 'NOTE' in name_upper:
        return 'Notes/Documentation'
    elif 'ACTIVITY' in name_upper or 'PROGRAM' in name_upper:
        return 'Activities/Programs'
    elif 'PAYMENT' in name_upper or 'RATES' in name_upper or 'RECEIVABLES' in name_upper or 'FINANCIAL' in name_upper:
        return 'Financial'
    elif 'SYNC' in name_upper or 'LOG' in name_upper:
        return 'Integration/Logging'
    elif 'TEMP' in name_upper or name_upper.startswith('TEMP_'):
        return 'Temporary Tables'
    elif '_OLD' in name_upper or '_BACKUP' in name_upper or '_BKP' in name_upper or 'EXCEPTIONS' in name_upper or '_HIST' in name_upper:
        return 'Archive/Historical'
    elif 'ISO_' in name_upper:
        return 'Reference Data - ISO'
    elif 'COMPANIES' in name_upper or 'DIVISION' in name_upper:
        return 'Configuration'
    else:
        return 'Other'

def main():
    # Directory containing HTML files (current directory by default)
    html_dir = '.'
    
    # Find all HTML files
    html_files = [f for f in os.listdir(html_dir) if f.endswith('.html') and f not in ['index.html', 'Tables.html']]
    
    print(f"Found {len(html_files)} HTML files to process...")
    
    all_tables = []
    all_columns = []
    
    for idx, html_file in enumerate(html_files, 1):
        try:
            print(f"Processing {idx}/{len(html_files)}: {html_file}")
            metadata = extract_table_metadata(html_file)
            
            # Add category
            category = categorize_table(metadata['table_name'])
            
            # Prepare table-level data
            table_info = {
                'Table Number': idx,
                'Table Name': metadata['table_name'],
                'Category': category,
                'Row Count': metadata['row_count'],
                'Column Count': metadata['column_count'],
                'Primary Keys': metadata['primary_keys'],
                'Indexed Columns': metadata['indexed_columns'],
                'Index Count': len(metadata['indexes']),
                'Foreign Key Count': len(metadata['foreign_keys']),
                'Description': metadata['description']
            }
            all_tables.append(table_info)
            
            # Prepare column-level data
            for col in metadata['columns']:
                col_info = {
                    'Table Name': metadata['table_name'],
                    'Column Name': col['column_name'],
                    'Data Type': col['data_type'],
                    'Max Length': col['max_length'],
                    'Nullable': col['nullable'],
                    'Is Identity': col['is_identity'],
                    'Default Value': col['default_value'],
                    'Is Primary Key': 'Yes' if col['is_primary_key'] else 'No',
                    'Is Indexed': 'Yes' if col['is_indexed'] else 'No'
                }
                all_columns.append(col_info)
                
        except Exception as e:
            print(f"Error processing {html_file}: {e}")
    
    # Create DataFrames
    df_tables = pd.DataFrame(all_tables)
    df_columns = pd.DataFrame(all_columns)
    
    # Sort by category and table name
    df_tables = df_tables.sort_values(['Category', 'Table Name'])
    df_tables['Table Number'] = range(1, len(df_tables) + 1)
    
    # Create summary by category
    summary = df_tables.groupby('Category').agg({
        'Table Name': 'count',
        'Row Count': lambda x: x.apply(lambda v: int(v.replace(',', '')) if v.isdigit() or (isinstance(v, str) and v.replace(',', '').isdigit()) else 0).sum()
    }).rename(columns={'Table Name': 'Table Count', 'Row Count': 'Total Rows'})
    
    # Export to Excel with multiple sheets
    output_file = 'Database_Complete_Metadata_Analysis.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_tables.to_excel(writer, sheet_name='Tables Summary', index=False)
        df_columns.to_excel(writer, sheet_name='All Columns', index=False)
        summary.to_excel(writer, sheet_name='Category Summary')
        
        # Format the Excel file
        workbook = writer.book
        
        # Auto-adjust column widths
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"\n✅ Analysis complete!")
    print(f"📊 Total tables processed: {len(all_tables)}")
    print(f"📋 Total columns extracted: {len(all_columns)}")
    print(f"💾 Output saved to: {output_file}")
    print(f"\nCategory Breakdown:")
    print(summary.to_string())

if __name__ == "__main__":
    main()