import streamlit as st
import pandas as pd
import uuid
from datetime import datetime
import io
import zipfile

# Set page config
st.set_page_config(
    page_title="Emerald Inventory - Excel to CSV Converter",
    page_icon="💎",
    layout="wide"
)

class ExcelToCSVProcessor:
    def __init__(self):
        # Define the 4 processing sheets and their stage mappings
        self.sheet_mapping = {
            'CUT SHEET': 'CUT',
            'GHAT SHEET': 'GHAT', 
            'MM SHEET': 'MM',
            'POLISH SHEET': 'POLISH'
        }
        
        # Results tracking
        self.results = {
            'lots_data': [],
            'processing_records_data': [],
            'sheets_processed': [],
            'errors': [],
            'validation_issues': []
        }
        
        # Track unique lots to avoid duplicates
        self.lots_dict = {}

    def normalize_date(self, date_value):
        """Normalize different date formats to YYYY-MM-DD"""
        if pd.isna(date_value) or date_value == "" or date_value is None:
            return None
        
        try:
            # If it's already a datetime object
            if isinstance(date_value, datetime):
                return date_value.strftime('%Y-%m-%d')
            
            # If it's a string, try different formats
            if isinstance(date_value, str):
                date_value = str(date_value).strip()
                
                # Try DD-MM-YYYY format (like 14-06-2025)
                if '-' in date_value and len(date_value.split('-')) == 3:
                    parts = date_value.split('-')
                    if len(parts[0]) == 2:  # DD-MM-YYYY
                        day, month, year = parts
                        return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                
                # Try DD/MM/YYYY format
                if '/' in date_value and len(date_value.split('/')) == 3:
                    parts = date_value.split('/')
                    if len(parts[0]) == 2:  # DD/MM/YYYY
                        day, month, year = parts
                        return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                
                # Try parsing as standard date string
                parsed_date = pd.to_datetime(date_value, errors='coerce')
                if not pd.isna(parsed_date):
                    return parsed_date.strftime('%Y-%m-%d')
            
            # Try pandas auto-parsing
            parsed_date = pd.to_datetime(date_value, errors='coerce')
            if not pd.isna(parsed_date):
                return parsed_date.strftime('%Y-%m-%d')
                
        except Exception as e:
            st.warning(f"Could not parse date {date_value}: {e}")
            
        return None

    def normalize_lot_number(self, lot_value):
        """Normalize lot number to string"""
        if pd.isna(lot_value) or lot_value == "" or lot_value is None:
            return None
        return str(lot_value).strip()

    def normalize_numeric(self, value, default=0):
        """Normalize numeric values"""
        if pd.isna(value) or value == "" or value is None:
            return default
        try:
            # Convert to float first, then ensure it's a clean number
            num_val = float(value)
            # Remove unnecessary decimal places for whole numbers
            if num_val.is_integer():
                return int(num_val)
            return round(num_val, 4)  # Round to 4 decimal places
        except:
            return default

    def normalize_integer(self, value, default=None):
        """Normalize integer values - removes decimals properly"""
        if pd.isna(value) or value == "" or value is None:
            return default
        try:
            # Convert to float first to handle "486.0" format, then to int
            float_val = float(value)
            return int(float_val)
        except:
            return default

    def add_lot(self, lot_number, lot_weight):
        """Add lot to lots dictionary if not exists"""
        if lot_number not in self.lots_dict:
            lot_id = str(uuid.uuid4())
            self.lots_dict[lot_number] = {
                'lot_id': lot_id,
                'lot_number': lot_number,
                'lot_weight': float(lot_weight),
                'status': 'active'
            }
            self.results['lots_data'].append(self.lots_dict[lot_number])
            return lot_id
        else:
            # Update weight if different
            existing_lot = self.lots_dict[lot_number]
            if float(existing_lot['lot_weight']) != float(lot_weight):
                existing_lot['lot_weight'] = float(lot_weight)
                # Update in results as well
                for lot in self.results['lots_data']:
                    if lot['lot_number'] == lot_number:
                        lot['lot_weight'] = float(lot_weight)
                        break
            return existing_lot['lot_id']

    def process_sheet(self, df, stage):
        """Process individual sheet data"""
        st.info(f"Processing {stage} sheet with {len(df)} rows")
        
        processed_count = 0
        errors_in_sheet = 0
        
        for index, row in df.iterrows():
            try:
                # Skip empty rows
                if pd.isna(row.iloc[0]) or row.iloc[0] == "":
                    continue
                
                # Extract data with proper column mapping
                process_date = self.normalize_date(row.iloc[0])  # DATE column
                lot_number = self.normalize_lot_number(row.iloc[1])  # LOT NO. column
                lot_weight = self.normalize_numeric(row.iloc[2], 0)  # LOT WEIGHT column
                
                if not process_date or not lot_number:
                    continue
                
                # Handle stage-specific data extraction
                if stage == 'CUT':
                    # CUT SHEET has no given pieces/weight data
                    given_pieces = None
                    given_weight = 0.0
                    received_pieces = self.normalize_integer(row.iloc[5] if len(row) > 5 else None)  # REC.P
                    received_weight = self.normalize_numeric(row.iloc[6] if len(row) > 6 else 0, 0)  # REC.W
                else:
                    # Other sheets have complete data
                    given_pieces = self.normalize_integer(row.iloc[3] if len(row) > 3 else None)  # GIVEN P.
                    given_weight = self.normalize_numeric(row.iloc[4] if len(row) > 4 else 0, 0)  # GIVEN W.
                    received_pieces = self.normalize_integer(row.iloc[5] if len(row) > 5 else None)  # REC.P
                    received_weight = self.normalize_numeric(row.iloc[6] if len(row) > 6 else 0, 0)  # REC.W
                
                # Add lot and get lot_id
                lot_id = self.add_lot(lot_number, lot_weight)
                
                # Create processing record
                processing_record = {
                    'record_id': str(uuid.uuid4()),
                    'lot_id': lot_id,
                    'stage': stage,
                    'process_date': process_date,
                    'given_pieces': given_pieces,
                    'given_weight': given_weight,
                    'received_pieces': received_pieces,
                    'received_weight': received_weight
                }
                
                self.results['processing_records_data'].append(processing_record)
                processed_count += 1
                
            except Exception as e:
                error_msg = f"Error processing row {index + 2} in {stage}: {str(e)}"
                self.results['errors'].append(error_msg)
                errors_in_sheet += 1
        
        st.success(f"Successfully processed {processed_count} rows from {stage} sheet")
        if errors_in_sheet > 0:
            st.warning(f"Encountered {errors_in_sheet} errors in {stage} sheet")
        
        self.results['sheets_processed'].append(stage)

    def validate_data_integrity(self):
        """Validate that all processing records have corresponding lots"""
        st.info("🔍 Validating data integrity...")
        
        # Get all lot_ids from lots
        lot_ids_in_lots = set(lot['lot_id'] for lot in self.results['lots_data'])
        
        # Get all lot_ids from processing records
        lot_ids_in_processing = set(record['lot_id'] for record in self.results['processing_records_data'])
        
        # Find orphaned processing records
        orphaned_lot_ids = lot_ids_in_processing - lot_ids_in_lots
        
        if orphaned_lot_ids:
            self.results['validation_issues'].append(
                f"Found {len(orphaned_lot_ids)} processing records with lot_ids not in lots table: {list(orphaned_lot_ids)[:5]}..."
            )
            return False
        
        # Find unused lots
        unused_lot_ids = lot_ids_in_lots - lot_ids_in_processing
        if unused_lot_ids:
            self.results['validation_issues'].append(
                f"Found {len(unused_lot_ids)} lots with no processing records"
            )
        
        st.success("✅ Data integrity validation passed!")
        return True

    def process_excel_file(self, uploaded_file):
        """Main method to process the entire Excel file"""
        st.info("Starting Excel processing...")
        
        try:
            # Read all sheets
            excel_file = pd.ExcelFile(uploaded_file)
            st.info(f"Found sheets: {excel_file.sheet_names}")
            
            # Process each of the 4 main processing sheets
            for sheet_name, stage in self.sheet_mapping.items():
                if sheet_name in excel_file.sheet_names:
                    st.subheader(f"Processing {sheet_name} → {stage}")
                    
                    # Read sheet data (skip header row)
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
                    
                    # Remove completely empty rows
                    df = df.dropna(how='all')
                    
                    # Show preview
                    with st.expander(f"Preview {sheet_name} data"):
                        st.dataframe(df.head())
                    
                    # Process the sheet
                    self.process_sheet(df, stage)
                else:
                    st.warning(f"Sheet '{sheet_name}' not found in Excel file")
            
            # Validate data integrity
            if not self.validate_data_integrity():
                st.error("❌ Data integrity validation failed! Check validation issues below.")
                return False
            
            st.success("Processing completed successfully!")
            return True
            
        except Exception as e:
            error_msg = f"Fatal error processing Excel file: {str(e)}"
            st.error(error_msg)
            self.results['errors'].append(error_msg)
            return False

    def generate_csv_files(self):
        """Generate CSV files for download"""
        csv_files = {}
        
        # Generate lots CSV
        if self.results['lots_data']:
            lots_df = pd.DataFrame(self.results['lots_data'])
            # Ensure proper data types for lots
            lots_df['lot_weight'] = lots_df['lot_weight'].astype(float)
            # Sort by lot_number for easier reference
            lots_df = lots_df.sort_values('lot_number')
            csv_files['lots.csv'] = lots_df.to_csv(index=False, float_format='%.4f')
        
        # Generate processing records CSV
        if self.results['processing_records_data']:
            processing_df = pd.DataFrame(self.results['processing_records_data'])
            
            # Ensure proper data types for processing records
            processing_df['given_weight'] = processing_df['given_weight'].astype(float)
            processing_df['received_weight'] = processing_df['received_weight'].astype(float)
            
            # Handle integer columns properly - convert None to empty string for CSV
            processing_df['given_pieces'] = processing_df['given_pieces'].apply(
                lambda x: '' if pd.isna(x) or x is None else int(x)
            )
            processing_df['received_pieces'] = processing_df['received_pieces'].apply(
                lambda x: '' if pd.isna(x) or x is None else int(x)
            )
            
            # Sort by process_date and stage for better organization
            processing_df = processing_df.sort_values(['process_date', 'stage'])
            csv_files['processing_records.csv'] = processing_df.to_csv(index=False, float_format='%.4f')
            
            # Also generate separate CSV for each stage
            for stage in self.results['sheets_processed']:
                stage_df = processing_df[processing_df['stage'] == stage].copy()
                csv_files[f'{stage.lower()}_records.csv'] = stage_df.to_csv(index=False, float_format='%.4f')
        
        return csv_files

def main():
    st.title("💎 Emerald Inventory - Excel to CSV Converter")
    st.markdown("Upload your Excel file to convert it into CSV files for Supabase import")
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Choose your Excel file",
        type=['xlsx', 'xls'],
        help="Upload the Excel file containing CUT SHEET, GHAT SHEET, MM SHEET, and POLISH SHEET"
    )
    
    if uploaded_file is not None:
        # Show file details
        st.info(f"📄 File: {uploaded_file.name} ({uploaded_file.size} bytes)")
        
        # Initialize processor
        processor = ExcelToCSVProcessor()
        
        # Process the file
        if processor.process_excel_file(uploaded_file):
            # Show results summary
            st.header("📊 Processing Results")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Unique Lots", len(processor.results['lots_data']))
            with col2:
                st.metric("Processing Records", len(processor.results['processing_records_data']))
            with col3:
                st.metric("Sheets Processed", len(processor.results['sheets_processed']))
            
            # Show validation issues if any
            if processor.results['validation_issues']:
                st.warning(f"⚠️ {len(processor.results['validation_issues'])} validation issues found:")
                with st.expander("View Validation Issues"):
                    for issue in processor.results['validation_issues']:
                        st.text(f"• {issue}")
            
            # Show errors if any
            if processor.results['errors']:
                st.warning(f"⚠️ {len(processor.results['errors'])} errors encountered:")
                with st.expander("View Errors"):
                    for error in processor.results['errors']:
                        st.text(f"• {error}")
            
            # Generate CSV files
            csv_files = processor.generate_csv_files()
            
            if csv_files:
                st.header("📥 Download CSV Files")
                
                # Add prominent import order warning
                st.error("""
                🚨 **CRITICAL: Import Order Matters!**
                
                **Step 1:** Download and import `lots.csv` FIRST into Supabase
                **Step 2:** Only then import `processing_records.csv` 
                
                Importing in wrong order will cause foreign key constraint errors!
                """)
                
                st.markdown("Download the generated CSV files to import into Supabase:")
                
                # Create download buttons for individual files
                col1, col2 = st.columns(2)
                
                with col1:
                    if 'lots.csv' in csv_files:
                        st.download_button(
                            label="🥇 Step 1: Download Lots CSV (Import FIRST)",
                            data=csv_files['lots.csv'],
                            file_name='lots.csv',
                            mime='text/csv',
                            type="primary"
                        )
                    
                    # Individual stage files
                    for stage in processor.results['sheets_processed']:
                        filename = f'{stage.lower()}_records.csv'
                        if filename in csv_files:
                            st.download_button(
                                label=f"📊 Download {stage} Records CSV",
                                data=csv_files[filename],
                                file_name=filename,
                                mime='text/csv'
                            )
                
                with col2:
                    if 'processing_records.csv' in csv_files:
                        st.download_button(
                            label="🥈 Step 2: Download All Processing Records CSV (Import SECOND)",
                            data=csv_files['processing_records.csv'],
                            file_name='processing_records.csv',
                            mime='text/csv',
                            type="secondary"
                        )
                
                # Create zip file with all CSVs
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for filename, content in csv_files.items():
                        zip_file.writestr(filename, content)
                
                zip_buffer.seek(0)
                
                st.download_button(
                    label="📦 Download All CSV Files (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name='emerald_inventory_csvs.zip',
                    mime='application/zip'
                )
                
                # Show preview of generated data
                st.header("👀 Data Preview")
                
                tab1, tab2 = st.tabs(["Lots Data", "Processing Records Data"])
                
                with tab1:
                    if processor.results['lots_data']:
                        lots_df = pd.DataFrame(processor.results['lots_data'])
                        st.dataframe(lots_df, use_container_width=True)
                        st.info(f"Total unique lots: {len(lots_df)}")
                    else:
                        st.info("No lots data generated")
                
                with tab2:
                    if processor.results['processing_records_data']:
                        processing_df = pd.DataFrame(processor.results['processing_records_data'])
                        st.dataframe(processing_df, use_container_width=True)
                        st.info(f"Total processing records: {len(processing_df)}")
                        
                        # Show breakdown by stage
                        stage_counts = processing_df['stage'].value_counts()
                        st.subheader("Records by Stage:")
                        for stage, count in stage_counts.items():
                            st.text(f"• {stage}: {count} records")
                    else:
                        st.info("No processing records data generated")
        
        # Enhanced instructions for Supabase import
        st.header("📚 Supabase Import Instructions")
        st.markdown("""
        ## 🚨 Critical Import Steps (Follow Exactly!)
        
        ### Step 1: Import Lots Table FIRST
        1. Go to Supabase → Table Editor → `lots` table
        2. Click "Insert" → "Import data from CSV"
        3. Upload `lots.csv`
        4. Settings: ✅ First row contains headers, ✅ Auto-detect data types
        5. **Wait for completion** before proceeding
        
        ### Step 2: Import Processing Records SECOND
        1. Go to Supabase → Table Editor → `processing_records` table
        2. Click "Insert" → "Import data from CSV"  
        3. Upload `processing_records.csv`
        4. Settings: ✅ First row contains headers, ✅ Auto-detect data types
        
        ## ⚠️ Troubleshooting Foreign Key Errors
        
        **If you get the foreign key constraint error:**
        1. **Clear both tables** in Supabase (Delete all rows)
        2. **Re-download** the CSV files from this tool
        3. **Import lots.csv FIRST**, wait for completion
        4. **Then import processing_records.csv**
        
        **Alternative Approach:**
        - Import individual stage CSV files instead of the combined processing_records.csv
        - This can help isolate any problematic records
        
        ## 📋 Data Validation Checks
        - ✅ All processing records have valid lot_id references
        - ✅ No orphaned records
        - ✅ Data types are properly formatted
        - ✅ Dates are in YYYY-MM-DD format
        """)

if __name__ == "__main__":
    main()
