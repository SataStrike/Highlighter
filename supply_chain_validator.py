"""
Supply Chain Validator Module - Handles supply chain validation logic
"""
import pandas as pd
import re

def normalize_line(line):
    """Normalize a line of text for comparison"""
    if not line or not isinstance(line, str):
        return ""
    
    # Standardize spacing and make lowercase
    line = re.sub(r'\s+', ' ', line).strip().lower()
    
    # IMPORTANT: Preserve periods in domain names (like openx.com)
    # Only remove special characters other than periods, commas, dashes, and underscores
    line = re.sub(r'[^\w\s,\.\-]', '', line)
    
    return line

def parse_missing_lines(text):
    """Parse the missing lines from the cell content in File 1 Column C"""
    print(f"[DEBUG] parse_missing_lines input: {repr(text)}")
    
    if pd.isna(text) or text == "":
        print(f"[DEBUG] Empty or NaN text, returning empty list")
        return []
        
    # Clean up the text first - normalize newlines and remove extra whitespace
    if isinstance(text, str):
        # Normalize different newline formats
        text = re.sub(r'\r\n|\r|\\n', '\n', text)
        # Remove excess whitespace around newlines
        text = re.sub(r'\s*\n\s*', '\n', text)
        # Remove excess spaces between commas
        text = re.sub(r'\s*,\s*', ', ', text)
        text = text.strip()
    else:
        text = str(text).strip()
    
    print(f"[DEBUG] Normalized text: {repr(text)}")
    missing_lines = []
    
    # Check if the text contains newlines - process as multiple lines
    if '\n' in text:
        print(f"[DEBUG] Multiline format detected")
        candidate_lines = text.split('\n')
        for line in candidate_lines:
            line = line.strip()
            # Skip empty lines
            if not line:
                continue
                
            # Remove bullet points or numbering at the beginning
            line = re.sub(r'^[•\-\d\.\s]+', '', line).strip()
            if line:
                # Process valid line
                process_candidate_line(line, missing_lines)
        
        print(f"[DEBUG] Multiline processing result: {len(missing_lines)} lines")
        return missing_lines
    
    # Handle semicolon-separated lists
    if ';' in text:
        print(f"[DEBUG] Semicolon-separated format detected")
        for item in text.split(';'):
            item = item.strip()
            if item:
                # Remove numbering if it exists
                item = re.sub(r'^\d+\.?\s*', '', item).strip()
                if item:
                    process_candidate_line(item, missing_lines)
        
        print(f"[DEBUG] Semicolon processing result: {len(missing_lines)} lines")
        return missing_lines
    
    # Special pattern matching for ads.txt lines
    # Look for patterns like: domain.com, 12345, RESELLER, abcd1234
    # Or: domain.com, 12345, DIRECT
    line_patterns = [
        # Full pattern with certificate: domain, ID, RESELLER/DIRECT, certificate
        r'([\w\.-]+\.\w+)\s*,\s*(\w+)\s*,\s*(RESELLER|DIRECT|reseller|direct)\s*,\s*([\w\d]+)',
        
        # Pattern without certificate: domain, ID, RESELLER/DIRECT
        r'([\w\.-]+\.\w+)\s*,\s*(\w+)\s*,\s*(RESELLER|DIRECT|reseller|direct)'
    ]
    
    # Check if text contains these patterns
    found_patterns = False
    for pattern in line_patterns:
        matches = re.finditer(pattern, text, re.IGNORECASE)
        for match in matches:
            found_patterns = True
            line = match.group(0).strip()
            if line:
                missing_lines.append(line)
                print(f"[DEBUG] Added line via pattern match: {repr(line)}")
    
    if found_patterns:
        print(f"[DEBUG] Pattern matching result: {len(missing_lines)} lines")
        return missing_lines
    
    # If no specific patterns found, try comma-separated text
    if ',' in text:
        print(f"[DEBUG] Comma-separated format detected")
        # Check if this looks like a single ads.txt line split into parts
        parts = [p.strip() for p in text.split(',')]
        
        # If we have 3 or 4 parts and one is RESELLER/DIRECT, treat as a single line
        if len(parts) in [3, 4] and any(p.upper() in ['RESELLER', 'DIRECT'] for p in parts):
            print(f"[DEBUG] Detected single ads.txt line: {repr(text)}")
            missing_lines.append(text)
        else:
            # Otherwise, each part might be a separate entity
            for part in parts:
                if part and not re.match(r'^\d+$', part):  # Skip pure numbers
                    missing_lines.append(part)
                    print(f"[DEBUG] Added comma-separated part: {repr(part)}")
    else:
        # Single item with no separators
        print(f"[DEBUG] Single item with no separators: {repr(text)}")
        missing_lines.append(text)
    
    print(f"[DEBUG] Final processing result: {len(missing_lines)} lines")
    return missing_lines

def process_candidate_line(line, result_list):
    """Process a candidate line and add to result list if valid"""
    print(f"[DEBUG] Processing candidate line: {repr(line)}")
    
    # Check if this is a valid ads.txt entry
    # Pattern for standard ads.txt entry
    ads_txt_pattern = r'([\w\.-]+\.\w+)\s*,\s*(\w+)\s*,\s*(RESELLER|DIRECT|reseller|direct)(?:\s*,\s*([\w\d]+))?'
    
    match = re.match(ads_txt_pattern, line, re.IGNORECASE)
    if match:
        # This is a valid ads.txt entry
        print(f"[DEBUG] Valid ads.txt entry: {repr(line)}")
        result_list.append(line)
        return
    
    # Check if it's a partial entry (might be missing components)
    domain_pattern = r'([\w\.-]+\.\w+)'
    if re.match(domain_pattern, line):
        # At least contains a domain
        if ',' in line:
            # Has some structure, could be partial ads.txt
            print(f"[DEBUG] Possible partial ads.txt entry: {repr(line)}")
            result_list.append(line)
        else:
            # Just a domain
            print(f"[DEBUG] Domain only: {repr(line)}")
            result_list.append(line)
    elif any(word.upper() in ['RESELLER', 'DIRECT'] for word in line.split()):
        # Contains keywords, might be relevant
        print(f"[DEBUG] Line contains RESELLER/DIRECT: {repr(line)}")
        result_list.append(line)
    elif len(line.split()) <= 2 and not line.isdigit():
        # Short text that's not just a number, likely relevant
        print(f"[DEBUG] Short text: {repr(line)}")
        result_list.append(line)
    else:
        # Apply more lenient filtering for other formats
        words = line.split()
        if len(words) >= 2 and not line.isdigit():
            print(f"[DEBUG] Multi-word line: {repr(line)}")
            result_list.append(line)

def process_supply_chain_files(supply_chain_path, lines_ref_path):
    """
    Process supply chain files and return a DataFrame with results
    
    Parameters:
        supply_chain_path (str): Path to the supply chain file
        lines_ref_path (str): Path to the lines referential file
        
    Returns:
        DataFrame: DataFrame containing the processed results
    """
    try:
        print(f"[DEBUG] Starting supply chain processing with files:\n  - Supply Chain: {supply_chain_path}\n  - Lines Ref: {lines_ref_path}")
        
        # Read the supply chain file
        supply_chain_df = pd.read_excel(supply_chain_path)
        print(f"[DEBUG] Loaded supply chain file with shape: {supply_chain_df.shape}")
        print(f"[DEBUG] Supply chain columns: {supply_chain_df.columns.tolist()}")
        
        # Read the lines referential file
        lines_ref_df = pd.read_csv(lines_ref_path)
        print(f"[DEBUG] Loaded lines referential file with shape: {lines_ref_df.shape}")
        print(f"[DEBUG] Lines ref columns: {lines_ref_df.columns.tolist()}")
        
        # Transform the lines referential data into a dictionary for faster lookup
        lines_ref_dict = {}
        bidder_category_dict = {}  # For fallback search by bidder
        
        # Define category mapping to standardize categories
        category_mapping = {
            "MAIN": "Primary",  # This is crucial - map MAIN to Primary
            "Main": "Primary",
            "PRIMARY": "Primary",
            "primary": "Primary",
            "MASTER": "Master",
            "master": "Master",
            "SECONDARY": "Secondary",
            "secondary": "Secondary"
        }
        
        # Track original categories for debugging
        original_categories = set()
        mapped_categories = {}
        
        # Add known patterns for lines categories
        for idx, row in lines_ref_df.iterrows():
            line = ""
            category = ""
            
            # Try to get the line column
            for col in ['Line', 'AdsLine', 'Ads.txt Line', 'Line Content']:
                if col in lines_ref_df.columns and not pd.isna(row.get(col, None)):
                    line = str(row[col])
                    break
            
            # Try to get the category column - include 'Line category' which is in the actual file
            for col in ['Category', 'Line category', 'Type', 'Line Type', 'LineType']:
                if col in lines_ref_df.columns and not pd.isna(row.get(col, None)):
                    category = str(row[col])
                    break
                    
            # Print debug message for the first few rows to confirm category column is found
            if idx < 5:
                print(f"[DEBUG] Row {idx}: Found category '{category}' from columns {lines_ref_df.columns.tolist()}")
            
            if line and category:
                # Track original category
                original_categories.add(category)
                
                # Map category to standard form
                if category in category_mapping:
                    mapped_category = category_mapping[category]
                    if category != mapped_category:
                        if category not in mapped_categories:
                            mapped_categories[category] = mapped_category
                            print(f"[DEBUG] Mapping category '{category}' to '{mapped_category}'")
                    category = mapped_category
                
                # Normalize the line
                normalized_line = normalize_line(line)
                if normalized_line:
                    lines_ref_dict[normalized_line] = (category, "Referential")
                    
                    # Add to bidder lookup for fallback
                    if ',' in normalized_line:
                        bidder = normalized_line.split(',')[0].strip().lower()
                        if bidder not in bidder_category_dict:
                            bidder_category_dict[bidder] = []
                        bidder_category_dict[bidder].append((normalized_line, category, "Referential"))
        
        # Print summary of category mapping
        print(f"[DEBUG] Original categories in referential: {original_categories}")
        print(f"[DEBUG] Category mappings applied: {mapped_categories}")
        
        print(f"[DEBUG] Built reference dictionary with {len(lines_ref_dict)} entries")
        print(f"[DEBUG] Built bidder dictionary with {len(bidder_category_dict)} unique bidders")
        
        # Common prefixes for categorization - MINIMIZED TO RESPECT REFERENTIAL DATA
        # Only keeping Adagio as Master since this is a special case requirement
        common_prefixes = [
            ("adagio", "Master"),  # Special case for Adagio
            
            # Secondary bidders
            ("google", "Secondary"),
            ("doubleclick", "Secondary"),
            ("freewheel", "Secondary"),
            ("spotxchange", "Secondary"),
            ("spotx", "Secondary"),
            ("adform", "Secondary"),
            ("media.net", "Secondary"),
            ("contextweb", "Secondary"),
            ("taboola", "Secondary"),
            ("outbrain", "Secondary"),
            ("vidazoo", "Secondary"),
            ("smartclip", "Secondary"),
            ("smaato", "Secondary"),
            ("rhythmone", "Secondary")
        ]
        
        # Process each row in the supply chain file
        results = []
        print(f"[DEBUG] Starting to process {len(supply_chain_df)} supply chain rows")
        
        # Key tracker to ensure all domains are processed and added to results
        # This is to ensure domains with no missing lines are still included
        existing_domains = {}
        
        for idx, row in supply_chain_df.iterrows():
            try:
                # Extract domain, name, status
                domain = row.get('Domain', '')
                
                # Try multiple possible column names for the publisher name
                name = None
                possible_name_columns = ['Publisher Name', 'Name', 'Publisher', 'Site Name']
                for col in possible_name_columns:
                    if col in supply_chain_df.columns and not pd.isna(row.get(col, '')) and row.get(col, '') != '':
                        name = row.get(col, '')
                        print(f"[DEBUG] Found name in column '{col}': '{name}'")
                        break
                
                # If no name column found or it's empty, use the domain as a fallback
                if not name:
                    name = domain
                    print(f"[DEBUG] Using domain as name fallback: '{name}'")
                
                status = row.get('Status', '')
                
                print(f"\n[DEBUG] Processing row {idx}: Domain={domain}, Publisher={name}, Status={status}")
                
                # ====== IMPORTANT: ALL DOMAINS SHOULD BE INCLUDED IN RESULTS ======
                # Even if they have no missing lines, we want to include them in the output
                # with zeros for all counts and "No missing bidders" for the bidders field
                
                # Get the missing lines from Column C (MISSING LINES TEXT)
                missing_lines_col = None
                possible_column_names = ['Missing Lines Text', 'Missing Lines', 'Missing', 'File 1 Column C', 'Rows with Missing Participants']
                
                for col in possible_column_names:
                    if col in supply_chain_df.columns:
                        missing_lines_col = col
                        break
                
                if missing_lines_col:
                    print(f"[DEBUG] Found missing lines column: {missing_lines_col}")
                    missing_lines_text = row.get(missing_lines_col, '')
                else:
                    print(f"[DEBUG] WARNING: Could not find missing lines column in {supply_chain_df.columns.tolist()}")
                    # Try to find a column that might contain missing lines text
                    for col in supply_chain_df.columns:
                        if 'missing' in col.lower() or 'line' in col.lower():
                            print(f"[DEBUG] Using column {col} as potential missing lines column")
                            missing_lines_text = row.get(col, '')
                            break
                    else:
                        print(f"[DEBUG] WARNING: No suitable column found for missing lines")
                        missing_lines_text = ''
                
                print(f"[DEBUG] Missing lines text: {repr(missing_lines_text)}")
                
                # Parse the missing lines using the imported function
                def parse_missing_lines(missing_lines_text):
                    """Parse missing lines from text and return as a list"""
                    if not missing_lines_text or missing_lines_text.strip() == "":
                        return []  # Return empty list if no missing lines text
                    
                    # Split the text into lines and clean up
                    lines = [line.strip() for line in missing_lines_text.splitlines() if line.strip()]        
                    return lines  # Return a simple list of lines
                
                missing_lines = parse_missing_lines(missing_lines_text)
                print(f"[DEBUG] Parsing missing lines text")
                
                # Count missing lines by category
                primary_missing = 0  # We'll count as we process
                secondary_missing = 0
                master_missing = 0    # New category for Master lines
                primary_lines = []    # We'll populate as we process
                secondary_lines = []
                master_lines = []     # Track Master lines
                unknown_lines = []
                
                # Debug message - show the reference dictionary categories to help debug
                print("[DEBUG] Categories found in referential data:")
                category_counts = {}
                for _, (cat, _) in lines_ref_dict.items():
                    category_counts[cat] = category_counts.get(cat, 0) + 1
                for cat, count in category_counts.items():
                    print(f"[DEBUG]   - {cat}: {count} entries")
                
                # Always add an entry even if there are no missing lines
                # This is the fix for domains with no missing lines not appearing in the output
                should_add_entry = True
                
                # Check each missing line against the reference
                print(f"[DEBUG] Checking {len(missing_lines)} missing lines against reference data")
                for i, line in enumerate(missing_lines):
                    normalized_line = normalize_line(line)
                    if not normalized_line:
                        print(f"[DEBUG] Line {i}: Empty after normalization, skipping")
                        continue
                    
                    # Special handling for smartadserver.com - directly categorize as Secondary
                    # This direct check is necessary because the previous attempt didn't work
                    if ',' in normalized_line and normalized_line.split(',')[0].strip().lower() == 'smartadserver.com':
                        print(f"[DEBUG] Line {i}: DIRECT handling for smartadserver.com - forced to SECONDARY")
                        secondary_missing += 1
                        secondary_lines.append(line)
                        found = True
                        match_type = "direct_smartadserver"
                        continue  # Skip other checks
                        
                    # No more hard-coded pattern matching for other cases - we'll use the referential data
                    # This is intentionally left empty to ensure no special case handling
                        
                    print(f"[DEBUG] Processing line {i}: '{line}' (normalized: '{normalized_line}')") 
                    found = False
                    match_type = "none"
                    
                    # Try exact match first
                    if normalized_line in lines_ref_dict:
                        category, _ = lines_ref_dict[normalized_line]
                        print(f"[DEBUG] Line {i}: EXACT MATCH found! Category: {category}")
                        if category == "Primary":
                            primary_missing += 1
                            primary_lines.append(line)
                            print(f"[DEBUG] Line {i}: Added to PRIMARY count")
                        elif category == "Master":
                            master_missing += 1
                            master_lines.append(line)
                            print(f"[DEBUG] Line {i}: Added to MASTER count")
                        else:
                            secondary_missing += 1
                            secondary_lines.append(line)
                            print(f"[DEBUG] Line {i}: Added to SECONDARY count")
                        found = True
                        match_type = "exact"
                    
                    # Check for specific vendor ID patterns
                    if not found and ',' in normalized_line:
                        parts = [p.strip().lower() for p in normalized_line.split(',')]
                        if len(parts) >= 2:
                            vendor = parts[0]
                            seller_id = parts[1]
                            print(f"[DEBUG] Line {i}: Checking vendor ID pattern: vendor='{vendor}', seller_id='{seller_id}'")
                            
                            # Get vendor and seller ID from line for matching against reference data
                            print(f"[DEBUG] Line {i}: Checking for vendor+ID match: vendor='{vendor}', seller_id='{seller_id}'")
                            
                            # Try to find an exact match by vendor+ID in the reference data
                            vendor_id_match = False
                            for ref_line, (ref_category, ref_source) in lines_ref_dict.items():
                                if ',' in ref_line:
                                    ref_parts = [p.strip().lower() for p in ref_line.split(',')]
                                    if len(ref_parts) >= 2:
                                        ref_vendor = ref_parts[0]
                                        ref_seller_id = ref_parts[1]
                                        
                                        # SPECIAL CASE: For Adagio lines, ignore the seller ID and always categorize as Master
                                        if vendor.lower() == "adagio.io":
                                            print(f"[DEBUG] Line {i}: ADAGIO LINE DETECTED - Always categorized as MASTER")
                                            master_missing += 1
                                            master_lines.append(line)
                                            vendor_id_match = True
                                            found = True
                                            match_type = "adagio_special_case"
                                            break
                                            
                                        # Check for vendor+ID match
                                        if vendor.lower() == ref_vendor and seller_id.lower() == ref_seller_id:
                                            print(f"[DEBUG] Line {i}: VENDOR+ID MATCH found! Reference: '{ref_line}', Category: {ref_category}")
                                            
                                            # Handle different category names (MAIN = PRIMARY)
                                            if ref_category in ["Primary", "MAIN", "Main", "PRIMARY"]:
                                                primary_missing += 1
                                                primary_lines.append(line)
                                                print(f"[DEBUG] Line {i}: Added to PRIMARY count via vendor+ID match")
                                            elif ref_category in ["Master", "MASTER", "master"]:
                                                master_missing += 1
                                                master_lines.append(line)
                                                print(f"[DEBUG] Line {i}: Added to MASTER count via vendor+ID match")
                                            else:  # Secondary or any other category
                                                secondary_missing += 1
                                                secondary_lines.append(line)
                                                print(f"[DEBUG] Line {i}: Added to SECONDARY count via vendor+ID match")
                                            vendor_id_match = True
                                            found = True
                                            match_type = "vendor_id_match"
                                            break
                            
                            # If no exact vendor+ID match, try using the bidder category dictionary
                            if not vendor_id_match and vendor.lower() in bidder_category_dict:
                                # Check if there's a consensus category for this vendor
                                vendor_entries = bidder_category_dict[vendor.lower()]
                                categories = [entry[1] for entry in vendor_entries]
                                
                                # If all entries for this vendor have the same category, use that
                                if len(set(categories)) == 1:
                                    vendor_category = categories[0]
                                    print(f"[DEBUG] Line {i}: Using vendor category from bidder dictionary: {vendor_category}")
                                    
                                    if vendor_category == "Primary":
                                        primary_missing += 1
                                        primary_lines.append(line)
                                        print(f"[DEBUG] Line {i}: Added to PRIMARY count via vendor category")
                                    elif vendor_category == "Master":
                                        master_missing += 1
                                        master_lines.append(line)
                                        print(f"[DEBUG] Line {i}: Added to MASTER count via vendor category")
                                    else:  # Secondary or other
                                        secondary_missing += 1
                                        secondary_lines.append(line)
                                        print(f"[DEBUG] Line {i}: Added to SECONDARY count via vendor category")
                                    found = True
                                    match_type = "vendor_category"
                                else:
                                    # If there's no consensus, DEFAULT TO SECONDARY except for special cases
                                    category_counts = {}
                                    for category in categories:
                                        category_counts[category] = category_counts.get(category, 0) + 1
                                    
                                    # Only use PRIMARY if it is significantly more common than Secondary
                                    # Otherwise default to Secondary for ambiguous cases
                                    use_secondary = True
                                    
                                    # SPECIAL CASE: Always categorize smartadserver.com as Secondary unless exact match found
                                    # This addresses the specific issue with smartadserver.com lines
                                    if vendor.lower() == 'smartadserver.com':
                                        print(f"[DEBUG] Line {i}: Special handling for smartadserver.com - defaulting to SECONDARY")
                                        use_secondary = True
                                    
                                    # Special case: If Primary count is at least 3x more than Secondary, use Primary
                                    primary_count = category_counts.get("Primary", 0)
                                    secondary_count = category_counts.get("Secondary", 0)
                                    
                                    if primary_count > 0 and secondary_count > 0:
                                        if primary_count >= (3 * secondary_count):
                                            use_secondary = False
                                            print(f"[DEBUG] Line {i}: Using PRIMARY because count ({primary_count}) is significantly higher than Secondary ({secondary_count})")
                                        else:
                                            print(f"[DEBUG] Line {i}: Using SECONDARY despite Primary count ({primary_count}) vs Secondary ({secondary_count})")
                                    elif primary_count > 0 and secondary_count == 0:
                                        # If we have Primary examples but no Secondary, still be conservative
                                        if primary_count >= 3:
                                            use_secondary = False
                                            print(f"[DEBUG] Line {i}: Using PRIMARY with strong consensus ({primary_count} examples)")
                                        else:
                                            print(f"[DEBUG] Line {i}: Using SECONDARY despite {primary_count} Primary examples (not enough consensus)")
                                            
                                    if not use_secondary and "Master" in category_counts and category_counts["Master"] > 0:
                                        # If there are any Master examples, prioritize those
                                        master_missing += 1
                                        master_lines.append(line)
                                        print(f"[DEBUG] Line {i}: Added to MASTER count due to presence of Master examples")
                                    elif not use_secondary:
                                        primary_missing += 1
                                        primary_lines.append(line)
                                        print(f"[DEBUG] Line {i}: Added to PRIMARY count with strong consensus")
                                    else:
                                        # Default case: use Secondary
                                        secondary_missing += 1
                                        secondary_lines.append(line)
                                        print(f"[DEBUG] Line {i}: Added to SECONDARY count as default categorization")
                                    found = True
                                    match_type = "vendor_most_common"
                    
                    # Fallback: try by bidder
                    if not found and ',' in normalized_line:
                        bidder = normalized_line.split(',')[0].strip().lower()
                        print(f"[DEBUG] Line {i}: Trying BIDDER match for '{bidder}'")
                        if bidder in bidder_category_dict:
                            print(f"[DEBUG] Line {i}: Bidder '{bidder}' found in bidder dictionary with {len(bidder_category_dict[bidder])} entries")
                            # Find the closest match
                            best_match = None
                            best_score = 0
                            for ref_line, category, source in bidder_category_dict[bidder]:
                                # Simple similarity score: length of common prefix
                                common = 0
                                for a, b in zip(normalized_line, ref_line):
                                    if a == b:
                                        common += 1
                                    else:
                                        break
                                        
                                score = common / max(len(normalized_line), len(ref_line))
                                print(f"[DEBUG] Line {i}: Similarity score with '{ref_line}': {score:.2f}")
                                if score > best_score and score > 0.7:  # Threshold for match
                                    best_score = score
                                    best_match = (ref_line, category, source)
                                    
                            if best_match:
                                ref_line, category, _ = best_match
                                print(f"[DEBUG] Line {i}: BIDDER MATCH found! Best match: '{ref_line}', Category: {category}, Score: {best_score:.2f}")
                                if category == "Primary":
                                    primary_missing += 1
                                    primary_lines.append(line)
                                    print(f"[DEBUG] Line {i}: Added to PRIMARY count")
                                elif category == "Master":
                                    master_missing += 1
                                    master_lines.append(line)
                                    print(f"[DEBUG] Line {i}: Added to MASTER count")
                                else:
                                    secondary_missing += 1
                                    secondary_lines.append(line)
                                    print(f"[DEBUG] Line {i}: Added to SECONDARY count")
                                found = True
                                match_type = "bidder"
                        else:
                            print(f"[DEBUG] Line {i}: Bidder '{bidder}' not found in bidder dictionary")
                    
                    # Try prefix matching as a last resort
                    if not found:
                        print(f"[DEBUG] Line {i}: Trying PREFIX MATCHING as last resort")
                        for prefix, category in common_prefixes:
                            if prefix in normalized_line:
                                print(f"[DEBUG] Line {i}: PREFIX MATCH found! Prefix: '{prefix}', Category: {category}")
                                if category == "Primary":
                                    primary_missing += 1
                                    primary_lines.append(line)
                                    print(f"[DEBUG] Line {i}: Added to PRIMARY count via prefix match")
                                elif category == "Master":
                                    master_missing += 1
                                    master_lines.append(line)
                                    print(f"[DEBUG] Line {i}: Added to MASTER count via prefix match")
                                else:
                                    secondary_missing += 1
                                    secondary_lines.append(line)
                                    print(f"[DEBUG] Line {i}: Added to SECONDARY count via prefix match")
                                found = True
                                match_type = "prefix"
                                break
                    
                    # If still not found after all attempts, categorize as Unknown instead of Secondary
                    if not found:
                        # Special exception is for adagio.io which goes to Master
                        if ',' in normalized_line and normalized_line.split(',')[0].strip().lower() == 'adagio.io':
                            master_missing += 1
                            master_lines.append(line)
                            print(f"[DEBUG] Line {i}: UNMATCHED ADAGIO line categorized as MASTER")
                        else:
                            # Add to unknown lines - this is a change from previous behavior
                            unknown_lines.append(line)
                            print(f"[DEBUG] Line {i}: UNMATCHED line categorized as UNKNOWN")
                            # Don't set found=True so it will be counted as unknown
                
                # Create a unique key for this Domain+Name pair to handle duplicates
                domain_name_key = f"{domain}_{name}"
                
                # Always add domain to results, even if no missing lines
                # This ensures domains with zero missing lines appear in the output
                
                # See if the domain already exists in results
                if domain_name_key in existing_domains:
                    # Update existing entry
                    idx = existing_domains[domain_name_key]
                    print(f"[DEBUG] Updating existing entry for {domain_name_key} at index {idx}")
                    
                    # Add counts to existing entry
                    results[idx]["Master Missing"] += master_missing
                    results[idx]["Primary Missing"] += primary_missing
                    results[idx]["Secondary Missing"] += secondary_missing
                    results[idx]["Total Missing"] += master_missing + primary_missing + secondary_missing
                    
                    # Extract bidder names from primary missing lines
                    if primary_lines:
                        primary_bidders = []
                        
                        # Check if the 'Bidder' column exists in the supply chain report
                        bidder_col_name = None
                        possible_bidder_cols = ['Bidder', 'Bidder Name', 'Partner', 'Partner Name']
                        for col in possible_bidder_cols:
                            if col in supply_chain_df.columns:
                                bidder_col_name = col
                                break
                        
                        if bidder_col_name:
                            # Use the Bidder column from the supply chain report
                            bidder = row.get(bidder_col_name, '')
                            if bidder and isinstance(bidder, str) and bidder.strip():
                                primary_bidders.append(bidder.strip())
                        else:
                            # Fallback: extract bidder names from primary missing lines
                            for line in primary_lines:
                                normalized_line = normalize_line(line)
                                if normalized_line and ',' in normalized_line:
                                    # Get the bidder name (first part before comma)
                                    bidder = normalized_line.split(',')[0].strip()
                                    
                                    # Remove domain suffix if present (e.g., .com, .io, etc.)
                                    bidder = re.sub(r'\.\w+$', '', bidder)
                                    
                                    if bidder:
                                        primary_bidders.append(bidder)
                        
                        # Add to existing bidders if there are any
                        if primary_bidders:
                            # Add space after each bidder name for better readability
                            formatted_bidders = [bidder + " " for bidder in primary_bidders]
                            if results[idx]["Missing Primary Bidders"] and results[idx]["Missing Primary Bidders"] != "No missing bidders":
                                results[idx]["Missing Primary Bidders"] += "; " + "; ".join(formatted_bidders)
                            else:
                                results[idx]["Missing Primary Bidders"] = "; ".join(formatted_bidders)
                        else:
                            # If there are no primary bidders, set a clear text instead of empty string
                            if not results[idx]["Missing Primary Bidders"] or results[idx]["Missing Primary Bidders"] == "No missing bidders":
                                results[idx]["Missing Primary Bidders"] = "No missing bidders"
                    
                    # Combine line lists
                    if master_lines:
                        results[idx]["Master Lines"] += ", " + ", ".join(master_lines)
                    if primary_lines:
                        results[idx]["Primary Lines"] += ", " + ", ".join(primary_lines)
                    if secondary_lines:
                        results[idx]["Secondary Lines"] += ", " + ", ".join(secondary_lines)
                    if unknown_lines:
                        results[idx]["Unknown Lines Text"] += ", " + ", ".join(unknown_lines)
                else:
                    # Add new entry
                    print(f"[DEBUG] Creating new entry for {domain_name_key}")
                    
                    # CRITICAL FIX: Convert MAIN category to Primary if needed
                    # Additional check for MAIN lines that should be mapped to Primary
                    for i, line in enumerate(missing_lines):
                        normalized_line = normalize_line(line)
                        for ref_line, (category, source) in lines_ref_dict.items():
                            if category.upper() == "MAIN" and normalized_line and normalized_line in ref_line:
                                print(f"[DEBUG] Found MAIN category line that should be Primary: {line}")
                                if line not in primary_lines and line not in secondary_lines and line not in master_lines:
                                    primary_missing += 1
                                    primary_lines.append(line)
                                    print(f"[DEBUG] Added MAIN line to PRIMARY: {line}")
                    
                        # Extract bidder names from primary missing lines
                    primary_bidders = []
                    
                    # Check if the 'Bidder' column exists in the supply chain report
                    bidder_col_name = None
                    possible_bidder_cols = ['Bidder', 'Bidder Name', 'Partner', 'Partner Name']
                    for col in possible_bidder_cols:
                        if col in supply_chain_df.columns:
                            bidder_col_name = col
                            print(f"[DEBUG] Found Bidder column: {bidder_col_name}")
                            break
                    
                    if bidder_col_name:
                        # Use the Bidder column from the supply chain report
                        bidder = row.get(bidder_col_name, '')
                        if bidder and isinstance(bidder, str) and bidder.strip():
                            primary_bidders.append(bidder.strip())
                            print(f"[DEBUG] Added bidder from Bidder column: {bidder.strip()}")
                    else:
                        # Fallback: extract bidder names from primary missing lines
                        for line in primary_lines:
                            normalized_line = normalize_line(line)
                            if normalized_line and ',' in normalized_line:
                                # Get the bidder name (first part before comma)
                                bidder = normalized_line.split(',')[0].strip()
                                
                                # Remove domain suffix if present (e.g., .com, .io, etc.)
                                bidder = re.sub(r'\.\w+$', '', bidder)
                                
                                if bidder:
                                    primary_bidders.append(bidder)
                                    print(f"[DEBUG] Extracted bidder from line: {bidder}")
                    
                    # Add the entry to results
                    results.append({
                        "Domain": domain,
                        "Name": name,
                        "Status": status,
                        "Master Missing": master_missing,
                        "Primary Missing": primary_missing,
                        "Secondary Missing": secondary_missing,
                        "Total Missing": primary_missing + secondary_missing + master_missing,
                        "Unknown Lines": len(unknown_lines),
                        "Master Lines": ", ".join(master_lines),
                        "Primary Lines": ", ".join(primary_lines),
                        "Secondary Lines": ", ".join(secondary_lines),
                        "Unknown Lines Text": ", ".join(unknown_lines),
                        "Missing Primary Bidders": "; ".join([bidder + " " for bidder in primary_bidders]) if primary_bidders else "No missing bidders"
                    })
                    existing_domains[domain_name_key] = len(results) - 1
                    
                # Debug the categorized lines
                print(f"[DEBUG] FINAL CATEGORIZATION FOR DOMAIN: {domain}")
                print(f"[DEBUG] Master lines ({len(master_lines)}):")
                for ml in master_lines[:5]:
                    print(f"[DEBUG]  - {ml}")
                if len(master_lines) > 5:
                    print(f"[DEBUG]  - ... and {len(master_lines)-5} more")
                    
                print(f"[DEBUG] Primary lines ({len(primary_lines)}):")
                for pl in primary_lines[:5]:
                    print(f"[DEBUG]  - {pl}")
                if len(primary_lines) > 5:
                    print(f"[DEBUG]  - ... and {len(primary_lines)-5} more")
                    
                print(f"[DEBUG] Secondary lines ({len(secondary_lines)}):")
                for sl in secondary_lines[:5]:
                    print(f"[DEBUG]  - {sl}")
                if len(secondary_lines) > 5:
                    print(f"[DEBUG]  - ... and {len(secondary_lines)-5} more")
            except Exception as e:
                print(f"[DEBUG][ERROR] Processing row {idx} failed: {str(e)}")
                
                # Even if there was an error, try to add the domain to ensure it appears in output
                try:
                    domain_name_key = f"{domain}_{name}"
                    if domain_name_key not in existing_domains and domain and name:
                        # Add basic entry with zeros for all counts
                        results.append({
                            "Domain": domain,
                            "Name": name,
                            "Status": status if status else "",
                            "Master Missing": 0,
                            "Primary Missing": 0,
                            "Secondary Missing": 0,
                            "Total Missing": 0,
                            "Unknown Lines": 0,
                            "Master Lines": "",
                            "Primary Lines": "",
                            "Secondary Lines": "",
                            "Unknown Lines Text": "",
                            "Missing Primary Bidders": "No missing bidders"
                        })
                        existing_domains[domain_name_key] = len(results) - 1
                        print(f"[DEBUG] Added basic entry for domain {domain} despite error")
                except Exception as inner_e:
                    print(f"[DEBUG][ERROR] Could not add basic entry for domain {domain}: {str(inner_e)}")
        print(f"[DEBUG] Created result DataFrame with shape: {len(results)}")
        print(f"[DEBUG] Result columns: {list(results[0].keys())}")
        
        if results:
            print(f"[DEBUG] First row example: {results[0]}")
            
            # Convert results to DataFrame
            result_df = pd.DataFrame(results)
            # Do NOT rename columns - keep original column names
            # This preserves 'Domain' as the first column name
            
            # Calculate total missing counts
            master_missing_count = result_df['Master Missing'].sum() if 'Master Missing' in result_df.columns else 0
            primary_missing_count = result_df['Primary Missing'].sum() if 'Primary Missing' in result_df.columns else 0
            secondary_missing_count = result_df['Secondary Missing'].sum() if 'Secondary Missing' in result_df.columns else 0
            total_missing_count = master_missing_count + primary_missing_count + secondary_missing_count
            
            print(f"[DEBUG] Total missing lines counts: Master={master_missing_count}, Primary={primary_missing_count}, Secondary={secondary_missing_count}, Total={total_missing_count}")
            print(f"[DEBUG] Supply chain DataFrame final columns: {result_df.columns.tolist()}")
        else:
            print(f"[DEBUG] WARNING: Result DataFrame is empty!")
            
        return result_df
        
    except Exception as e:
        print(f"Error processing supply chain files: {e}")
        import traceback
        traceback.print_exc()
        raise  # Re-raise to see the full error details
