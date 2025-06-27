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
        
        # Generated CSVs storage
        self.generated_lots_csv = None
        self.generated_lots_df = None

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

    def process_sheet_for_records(self, df, stage):
        """Process individual sheet data for processing records"""
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
                    'lot_id': None,  # Will be filled later by matching lot_number from lots.csv
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

    def process_excel_file_for_lots(self, uploaded_file):
        """STEP 1: Process Excel file and generate ONLY lots.csv"""
        st.info("🔍 STEP 1: Processing Excel file to generate lots.csv...")
        
        try:
            # Read all sheets
            excel_file = pd.ExcelFile(uploaded_file)
            st.info(f"Found sheets: {excel_file.sheet_names}")
            
            # Collect all unique lots from all sheets
            for sheet_name, stage in self.sheet_mapping.items():
                if sheet_name in excel_file.sheet_names:
                    st.info(f"Scanning {sheet_name} for unique lots...")
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
                    df = df.dropna(how='all')
                    self.collect_unique_lots(df)
                else:
                    st.warning(f"Sheet '{sheet_name}' not found in Excel file")
            
            # Generate lots data
            self.results['lots_data'] = list(self.unique_lots.values())
            st.success(f"✅ Found {len(self.unique_lots)} unique lots")
            
            # Generate lots.csv
            if self.results['lots_data']:
                lots_df = pd.DataFrame(self.results['lots_data'])
                lots_df['lot_weight'] = lots_df['lot_weight'].astype(float)
                lots_df = lots_df.sort_values('lot_number')
                
                # Store the generated CSV and DataFrame
                self.generated_lots_csv = lots_df.to_csv(index=False, float_format='%.4f')
                self.generated_lots_df = lots_df.copy()
                
                st.success(f"✅ Generated lots.csv with {len(lots_df)} unique lots")
                return True
            else:
                st.error("❌ No lots data found")
                return False
            
        except Exception as e:
            error_msg = f"Fatal error processing Excel file: {str(e)}"
            st.error(error_msg)
            self.results['errors'].append(error_msg)
            return False

    def process_excel_file_for_records(self, uploaded_file):
        """STEP 2: Process Excel file for processing records"""
        st.info("📊 STEP 2: Processing Excel file for processing records...")
        
        try:
            # Clear previous processing records
            self.results['processing_records_data'] = []
            self.results['sheets_processed'] = []
            
            # Read all sheets and process for records
            excel_file = pd.ExcelFile(uploaded_file)
            
            for sheet_name, stage in self.sheet_mapping.items():
                if sheet_name in excel_file.sheet_names:
                    st.subheader(f"Processing {sheet_name} → {stage}")
                    
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
                    df = df.dropna(how='all')
                    
                    # Show preview
                    with st.expander(f"Preview {sheet_name} data"):
                        st.dataframe(df.head())
                    
                    self.process_sheet_for_records(df, stage)
                else:
                    st.warning(f"Sheet '{sheet_name}' not found in Excel file")
            
            st.success("✅ Processing records data collected successfully!")
            return True
            
        except Exception as e:
            error_msg = f"Fatal error processing Excel file for records: {str(e)}"
            st.error(error_msg)
            self.results['errors'].append(error_msg)
            return False

    def generate_processing_records_csv(self):
        """Generate processing_records.csv using exact lot_ids from generated lots.csv"""
        if not self.generated_lots_df is not None:
            st.error("❌ Please generate lots.csv first!")
            return None
        
        if not self.results['processing_records_data']:
            st.error("❌ No processing records data found!")
            return None
        
        st.info("🔗 Matching lot_numbers with lot_ids from lots.csv...")
        
        # Create processing DataFrame
        processing_df = pd.DataFrame(self.results['processing_records_data'])
        
        # Create lot_number to lot_id mapping from the GENERATED lots.csv
        lot_mapping = {}
        for _, row in self.generated_lots_df.iterrows():
            lot_mapping[row['lot_number']] = row['lot_id']
        
        st.info(f"Available lot_ids in lots.csv: {len(lot_mapping)}")
        
        # Match lot_numbers and fill lot_ids
        matched_count = 0
        unmatched_count = 0
        unmatched_lots = []
        
        for idx, row in processing_df.iterrows():
            lot_number = row['lot_number']
            
            if lot_number in lot_mapping:
                # Use the EXACT lot_id from lots.csv
                processing_df.at[idx, 'lot_id'] = lot_mapping[lot_number]
                matched_count += 1
            else:
                st.error(f"❌ ERROR: lot_number '{lot_number}' not found in lots.csv")
                unmatched_lots.append(lot_number)
                unmatched_count += 1
        
        if unmatched_count > 0:
            st.error(f"❌ {unmatched_count} processing records could not be matched")
            st.error(f"Unmatched lot_numbers: {unmatched_lots}")
            return None
        
        st.success(f"✅ Successfully matched {matched_count} processing records with lot_ids from lots.csv")
        
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
        
        return processing_df.to_csv(index=False, float_format='%.4f')

# Initialize session state
if 'processor' not in st.session_state:
    st.session_state.processor = ExcelToCSVProcessor()
if 'lots_generated' not in st.session_state:
    st.session_state.lots_generated = False
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None

def main():
    st.title("💎 Emerald Inventory - Excel to CSV Converter")
    st.markdown("**Step-by-step approach:** Generate lots.csv first, then processing_records.csv with exact lot_id matching")
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Choose your Excel file",
        type=['xlsx', 'xls'],
        help="Upload the Excel file containing CUT SHEET, GHAT SHEET, MM SHEET, and POLISH SHEET"
    )
    
    if uploaded_file is not None:
        st.session_state.uploaded_file = uploaded_file
        st.info(f"📄 File: {uploaded_file.name} ({uploaded_file.size} bytes)")
        
        # STEP 1: Generate lots.csv
        st.header("🚀 STEP 1: Generate Lots CSV")
        
        if st.button("📊 Generate Lots CSV", type="primary"):
            st.session_state.processor = ExcelToCSVProcessor()  # Reset processor
            
            if st.session_state.processor.process_excel_file_for_lots(uploaded_file):
                st.session_state.lots_generated = True
                st.success("✅ Lots CSV generated successfully!")
            else:
                st.error("❌ Failed to generate lots CSV")
                st.session_state.lots_generated = False
        
        # Show lots.csv download if generated
        if st.session_state.lots_generated and st.session_state.processor.generated_lots_csv:
            st.subheader("📥 Download Lots CSV")
            
            # Show lots preview
            with st.expander("👀 Preview Lots Data"):
                st.dataframe(st.session_state.processor.generated_lots_df, use_container_width=True)
                st.info(f"Total unique lots: {len(st.session_state.processor.generated_lots_df)}")
            
            st.download_button(
                label="📊 Download lots.csv",
                data=st.session_state.processor.generated_lots_csv,
                file_name='lots.csv',
                mime='text/csv',
                type="primary"
            )
            
            st.success("✅ **lots.csv is ready!** Now generate processing_records.csv below.")
            
            # STEP 2: Generate processing_records.csv
            st.header("🚀 STEP 2: Generate Processing Records CSV")
            st.info("This will use the exact lot_ids from the lots.csv generated above")
            
            if st.button("📋 Generate Processing Records CSV", type="secondary"):
                if st.session_state.processor.process_excel_file_for_records(uploaded_file):
                    # Generate processing records CSV with exact lot_id matching
                    processing_csv = st.session_state.processor.generate_processing_records_csv()
                    
                    if processing_csv:
                        st.success("✅ Processing records CSV generated successfully!")
                        
                        # Show processing records preview
                        processing_df = pd.DataFrame(st.session_state.processor.results['processing_records_data'])
                        
                        with st.expander("👀 Preview Processing Records Data"):
                            st.dataframe(processing_df, use_container_width=True)
                            st.info(f"Total processing records: {len(processing_df)}")
                            
                            # Show breakdown by stage
                            stage_counts = processing_df['stage'].value_counts()
                            st.subheader("Records by Stage:")
                            for stage, count in stage_counts.items():
                                st.text(f"• {stage}: {count} records")
                        
                        # Download processing records CSV
                        st.download_button(
                            label="📋 Download processing_records.csv",
                            data=processing_csv,
                            file_name='processing_records.csv',
                            mime='text/csv',
                            type="primary"
                        )
                        
                        # Create zip with both files
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            zip_file.writestr('lots.csv', st.session_state.processor.generated_lots_csv)
                            zip_file.writestr('processing_records.csv', processing_csv)
                        
                        zip_buffer.seek(0)
                        
                        st.download_button(
                            label="📦 Download Both CSV Files (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name='emerald_inventory_csvs.zip',
                            mime='application/zip'
                        )
                        
                        st.success("""
                        ✅ **PERFECT MATCHING COMPLETED!**
                        
                        **What happened:**
                        1. 📊 Generated lots.csv with unique lot_id for each lot_number
                        2. 📋 Generated processing_records.csv using EXACT same lot_id from lots.csv
                        3. 🔗 Matched lot_number between both CSVs - NO random lot_ids!
                        
                        **Ready for Supabase import:**
                        - Import lots.csv first
                        - Then import processing_records.csv
                        - lot_ids will match perfectly!
                        """)
                    else:
                        st.error("❌ Failed to generate processing records CSV")
                else:
                    st.error("❌ Failed to process Excel file for records")
        
        # Show errors if any
        if st.session_state.processor.results['errors']:
            st.warning(f"⚠️ {len(st.session_state.processor.results['errors'])} errors encountered:")
            with st.expander("View Errors"):
                for error in st.session_state.processor.results['errors']:
                    st.text(f"• {error}")
    
    # Instructions
    st.header("📚 How This Works")
    st.markdown("""
    ## 🎯 **Step-by-Step Approach**
    
    ### Why This Method Works:
    1. **🔍 STEP 1:** Generate lots.csv first with unique lot_id for each lot_number
    2. **📋 STEP 2:** Generate processing_records.csv and match lot_number to use SAME lot_id from lots.csv
    3. **✅ Result:** Both CSVs have perfectly matching lot_ids - no randomness!
    
    ### 🚀 Supabase Import:
    1. Import `lots.csv` first
    2. Import `processing_records.csv` second
    3. All lot_ids will match perfectly!
    
    ### 🔍 Verification Query:
    ```sql
    -- Check perfect lot_id matching
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
    ORDER BY pr.lot_number;
    ```
    """)

if __name__ == "__main__":
    main()
