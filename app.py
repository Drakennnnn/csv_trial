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
            'errors': []
        }
        
        # Track unique lots - lot_number -> lot_data mapping
        self.unique_lots = {}

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
            num_val = float(value)
            if num_val.is_integer():
                return int(num_val)
            return round(num_val, 4)
        except:
            return default

    def normalize_integer(self, value, default=None):
        """Normalize integer values"""
        if pd.isna(value) or value == "" or value is None:
            return default
        try:
            float_val = float(value)
            return int(float_val)
        except:
            return default

    def collect_unique_lots(self, df):
        """Collect unique lots from a sheet"""
        for index, row in df.iterrows():
            try:
                # Skip empty rows
                if pd.isna(row.iloc[0]) or row.iloc[0] == "":
                    continue
                
                lot_number = self.normalize_lot_number(row.iloc[1])  # LOT NO. column
                lot_weight = self.normalize_numeric(row.iloc[2], 0)  # LOT WEIGHT column
                
                if lot_number:
                    # Store unique lots
                    if lot_number not in self.unique_lots:
                        self.unique_lots[lot_number] = {
                            'lot_id': str(uuid.uuid4()),
                            'lot_number': lot_number,
                            'lot_weight': float(lot_weight),
                            'status': 'active'
                        }
                    else:
                        # Update weight if different
                        if float(self.unique_lots[lot_number]['lot_weight']) != float(lot_weight):
                            self.unique_lots[lot_number]['lot_weight'] = float(lot_weight)
                            
            except Exception as e:
                error_msg = f"Error collecting lot from row {index + 2}: {str(e)}"
                self.results['errors'].append(error_msg)

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
                
                # Extract data - ALL SHEETS HAVE SAME COLUMNS
                process_date = self.normalize_date(row.iloc[0])  # DATE column
                lot_number = self.normalize_lot_number(row.iloc[1])  # LOT NO. column
                lot_weight = self.normalize_numeric(row.iloc[2], 0)  # LOT WEIGHT column
                given_pieces = self.normalize_integer(row.iloc[3] if len(row) > 3 else None)  # GIVEN P.
                given_weight = self.normalize_numeric(row.iloc[4] if len(row) > 4 else 0, 0)  # GIVEN W.
                received_pieces = self.normalize_integer(row.iloc[5] if len(row) > 5 else None)  # REC.P
                received_weight = self.normalize_numeric(row.iloc[6] if len(row) > 6 else 0, 0)  # REC.W
                
                if not process_date or not lot_number:
                    continue
                
                # Create processing record with lot_number (lot_id will be filled later)
                processing_record = {
                    'record_id': str(uuid.uuid4()),
                    'lot_id': None,  # Will be filled later by matching lot_number
                    'lot_number': lot_number,  # Keep for matching
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

    def process_excel_file(self, uploaded_file):
        """Main method to process the entire Excel file"""
        st.info("Starting Excel processing...")
        
        try:
            # Read all sheets
            excel_file = pd.ExcelFile(uploaded_file)
            st.info(f"Found sheets: {excel_file.sheet_names}")
            
            # STEP 1: First pass - collect all unique lots from all sheets
            st.info("🔍 Step 1: Collecting unique lots from all sheets...")
            for sheet_name, stage in self.sheet_mapping.items():
                if sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
                    df = df.dropna(how='all')
                    self.collect_unique_lots(df)
            
            # Generate lots data
            self.results['lots_data'] = list(self.unique_lots.values())
            st.success(f"✅ Found {len(self.unique_lots)} unique lots")
            
            # STEP 2: Second pass - process each sheet for processing records
            st.info("📊 Step 2: Processing sheets for processing records...")
            for sheet_name, stage in self.sheet_mapping.items():
                if sheet_name in excel_file.sheet_names:
                    st.subheader(f"Processing {sheet_name} → {stage}")
                    
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
                    df = df.dropna(how='all')
                    
                    # Show preview
                    with st.expander(f"Preview {sheet_name} data"):
                        st.dataframe(df.head())
                    
                    self.process_sheet(df, stage)
                else:
                    st.warning(f"Sheet '{sheet_name}' not found in Excel file")
            
            st.success("✅ Processing completed successfully!")
            return True
            
        except Exception as e:
            error_msg = f"Fatal error processing Excel file: {str(e)}"
            st.error(error_msg)
            self.results['errors'].append(error_msg)
            return False

    def generate_csv_files(self):
        """Generate CSV files for download"""
        csv_files = {}
        
        # STEP 1: Generate lots.csv FIRST
        if self.results['lots_data']:
            lots_df = pd.DataFrame(self.results['lots_data'])
            lots_df['lot_weight'] = lots_df['lot_weight'].astype(float)
            lots_df = lots_df.sort_values('lot_number')
            csv_files['lots.csv'] = lots_df.to_csv(index=False, float_format='%.4f')
            st.success(f"✅ Generated lots.csv with {len(lots_df)} unique lots")
        
        # STEP 2: Generate processing_records.csv with lot_id matching
        if self.results['processing_records_data']:
            processing_df = pd.DataFrame(self.results['processing_records_data'])
            
            # CRITICAL: Match lot_number and fill lot_id from unique_lots
            matched_count = 0
            unmatched_count = 0
            
            for idx, row in processing_df.iterrows():
                lot_number = row['lot_number']
                
                if lot_number in self.unique_lots:
                    # Use the SAME lot_id from lots data
                    processing_df.at[idx, 'lot_id'] = self.unique_lots[lot_number]['lot_id']
                    matched_count += 1
                else:
                    st.error(f"❌ ERROR: Could not find lot_id for lot_number: {lot_number}")
                    unmatched_count += 1
            
            st.success(f"✅ Matched {matched_count} processing records with lot_ids")
            if unmatched_count > 0:
                st.error(f"❌ {unmatched_count} processing records could not be matched")
            
            # Format columns properly
            processing_df['given_weight'] = processing_df['given_weight'].astype(float)
            processing_df['received_weight'] = processing_df['received_weight'].astype(float)
            
            # Handle integer columns
            processing_df['given_pieces'] = processing_df['given_pieces'].apply(
                lambda x: '' if pd.isna(x) or x is None else int(x)
            )
            processing_df['received_pieces'] = processing_df['received_pieces'].apply(
                lambda x: '' if pd.isna(x) or x is None else int(x)
            )
            
            # Column order
            column_order = ['record_id', 'lot_id', 'lot_number', 'stage', 'process_date', 
                          'given_pieces', 'given_weight', 'received_pieces', 'received_weight']
            processing_df = processing_df[column_order]
            
            # Sort by process_date and stage
            processing_df = processing_df.sort_values(['process_date', 'stage'])
            csv_files['processing_records.csv'] = processing_df.to_csv(index=False, float_format='%.4f')
            
            # Generate separate CSV for each stage
            for stage in self.results['sheets_processed']:
                stage_df = processing_df[processing_df['stage'] == stage].copy()
                csv_files[f'{stage.lower()}_records.csv'] = stage_df.to_csv(index=False, float_format='%.4f')
        
        return csv_files

def main():
    st.title("💎 Emerald Inventory - Excel to CSV Converter")
    st.markdown("Upload your Excel file to convert it into CSV files for direct Supabase import")
    
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
                
                st.success("""
                ✅ **PERFECT MATCHING IMPLEMENTED!**
                
                **How it works:**
                1. 🔍 **Step 1:** Scanned all sheets and collected unique lots with generated lot_id
                2. 📊 **Step 2:** Generated lots.csv with lot_id for each lot_number  
                3. 🔗 **Step 3:** Generated processing_records.csv and matched lot_number to use SAME lot_id from lots.csv
                4. ✅ **Result:** Both CSVs have matching lot_id for same lot_number - NO random IDs!
                """)
                
                st.markdown("Download the generated CSV files:")
                
                # Create download buttons
                col1, col2 = st.columns(2)
                
                with col1:
                    if 'lots.csv' in csv_files:
                        st.download_button(
                            label="📊 Download Lots CSV",
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
                            label="📋 Download Processing Records CSV",
                            data=csv_files['processing_records.csv'],
                            file_name='processing_records.csv',
                            mime='text/csv',
                            type="primary"
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
        
        # Instructions 
        st.header("📚 Import Instructions")
        st.markdown("""
        ## 🚀 Direct Import Process
        
        ### Step 1: Import Lots Table
        1. Go to Supabase → Table Editor → `lots` table
        2. Click "Insert" → "Import data from CSV"
        3. Upload `lots.csv`
        4. ✅ First row contains headers, ✅ Auto-detect data types
        
        ### Step 2: Import Processing Records Table
        1. Go to Supabase → Table Editor → `processing_records` table  
        2. Click "Insert" → "Import data from CSV"
        3. Upload `processing_records.csv`
        4. ✅ First row contains headers, ✅ Auto-detect data types
        
        ## ✅ What's Fixed Now
        - ✅ **Two-pass processing:** First pass collects unique lots, second pass processes records
        - ✅ **Same lot_id for same lot_number:** No random generation during processing
        - ✅ **Perfect matching:** processing_records.csv uses exact lot_id from lots.csv
        - ✅ **lot_number included:** For easy verification and debugging
        - ✅ **All sheets have same columns:** Simplified processing logic
        
        ## 🔍 Verification Query
        ```sql
        -- Verify lot_id matching between tables
        SELECT 
            pr.lot_number,
            pr.lot_id,
            l.lot_number as lots_table_lot_number,
            CASE 
                WHEN pr.lot_id = l.lot_id AND pr.lot_number = l.lot_number 
                THEN '✅ PERFECT MATCH' 
                ELSE '❌ MISMATCH' 
            END as match_status
        FROM processing_records pr
        LEFT JOIN lots l ON pr.lot_id = l.lot_id
        ORDER BY pr.lot_number
        LIMIT 20;
        ```
        """)

if __name__ == "__main__":
    main()
