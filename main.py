import re
from typing import List, Dict


def extract_age_ranges(text: str) -> Dict[str, Dict]:
    """
    Extract age ranges dynamically from policy text,
    with proper defaults if no match is found.
    """
    age_ranges = {
        "employee": {"min_years": 18, "min_months": 0, "max_years": 65, "max_months": 0},
        "spouse": {"min_years": 18, "min_months": 0, "max_years": 65, "max_months": 0},
        "children": {"min_years": 0, "min_months": 3, "max_years": 25, "max_months": 0},
        "parents": {"min_years": 40, "min_months": 0, "max_years": 100, "max_months": 0}
    }

    # Employee / Spouse
    employee_age_match = re.search(
        r'employee.*?(?:age|years?)\s*(?:between|from|range)?\s*(\d+)\s*(?:to|-)\s*(\d+)',
        text,
        re.IGNORECASE
    )
    if employee_age_match:
        min_age = int(employee_age_match.group(1))
        max_age = int(employee_age_match.group(2))
        age_ranges["employee"]["min_years"] = min_age
        age_ranges["employee"]["max_years"] = max_age
        age_ranges["spouse"]["min_years"] = min_age
        age_ranges["spouse"]["max_years"] = max_age

    # Children
    children_age_match = re.search(
        r'(?:children|dependent).*?(?:age|years?)\s*(?:between|from|range)?\s*(\d+)\s*(?:to|-)\s*(\d+)',
        text,
        re.IGNORECASE
    )
    if children_age_match:
        min_age = int(children_age_match.group(1))
        max_age = int(children_age_match.group(2))
        age_ranges["children"]["min_years"] = min_age
        age_ranges["children"]["max_years"] = max_age

    # over X days pattern for children minimum
    children_days_match = re.search(r'over\s+(\d+)\s+days?', text, re.IGNORECASE)
    if children_days_match:
        days = int(children_days_match.group(1))
        months = days // 30
        age_ranges["children"]["min_months"] = months

    # Parents
    parents_age_match = re.search(
        r'(?:parents?|father|mother).*?(?:age|years?)\s*(?:between|from|range)?\s*(\d+)\s*(?:to|-)\s*(\d+)',
        text,
        re.IGNORECASE
    )
    if parents_age_match:
        min_age = int(parents_age_match.group(1))
        max_age = int(parents_age_match.group(2))
        age_ranges["parents"]["min_years"] = min_age
        age_ranges["parents"]["max_years"] = max_age

    return age_ranges

def extract_sublimits_from_endorsement_5i(text: str) -> List[Dict]:
    """Extract sublimits from Endt. No. 5(i) - Room, Boarding Expenses and Intensive Care Unit"""
    sublimits = []
    
    # Look for Endt. No. 5(i) patterns
    endorsement_5i_match = re.search(r'Endt\.\s*No\.\s*5\(i\).*?(?=Endt\.\s*No\.|Endorsement\s*No\.|$)', text, re.IGNORECASE | re.DOTALL)
    if endorsement_5i_match:
        endorsement_5i_text = endorsement_5i_match.group(0)
        print(f"[DEBUG] Found Endt. No. 5(i) section")
        
        # Extract Room, Boarding Expenses - look for percentage pattern
        room_boarding_patterns = [
            r'Room,\s*Boarding\s+Expenses.*?(\d+)%\s+of\s+the\s+Sum\s+Insured\s+per\s+day',
            r'Room.*?Boarding.*?(\d+)%\s+of\s+the\s+Sum\s+Insured',
            r'Room.*?Boarding.*?(\d+)%',
            r'Room,\s*Boarding.*?(\d+)%'
        ]
        
        room_boarding_percentage = None
        for pattern in room_boarding_patterns:
            room_boarding_match = re.search(pattern, endorsement_5i_text, re.IGNORECASE | re.DOTALL)
            if room_boarding_match:
                room_boarding_percentage = int(room_boarding_match.group(1))
                break
        
        if room_boarding_percentage:
            # Calculate amount using percentage and sum insured of 500000
            calculated_amount = int((room_boarding_percentage / 100) * 500000)
            
            sublimits.append({
                "applicable": "Yes",
                "type": "Room, Boarding Expenses",
                "limit": str(calculated_amount)
            })
            print(f"[DEBUG] Found Room, Boarding Expenses: {room_boarding_percentage}% = {calculated_amount}")
        
        # Extract Intensive Care Unit - look for percentage pattern
        icu_patterns = [
            r'Intensive\s+Care\s+Unit.*?(\d+)%\s+of\s+the\s+Sum\s+Insured\s+per\s+day',
            r'Intensive\s+Care.*?(\d+)%\s+of\s+the\s+Sum\s+Insured',
            r'Intensive\s+Care.*?(\d+)%',
            r'ICU.*?(\d+)%',
            r'Intensive.*?(\d+)%'
        ]
        
        icu_percentage = None
        for pattern in icu_patterns:
            icu_match = re.search(pattern, endorsement_5i_text, re.IGNORECASE | re.DOTALL)
            if icu_match:
                icu_percentage = int(icu_match.group(1))
                break
        
        if icu_percentage:
            # Calculate amount using percentage and sum insured of 500000
            calculated_amount = int((icu_percentage / 100) * 500000)
            
            sublimits.append({
                "applicable": "Yes",
                "type": "Intensive Care Unit",
                "limit": str(calculated_amount)
            })
            print(f"[DEBUG] Found Intensive Care Unit: {icu_percentage}% = {calculated_amount}")
    
    print(f"[DEBUG] Total 5(i) sublimits extracted: {len(sublimits)}")
    for i, sublimit in enumerate(sublimits):
        print(f"[DEBUG] {i+1}. Type: '{sublimit['type']}', Limit: '{sublimit['limit']}'")
    
    return sublimits


def extract_all_sublimits_from_endorsement_5ii(text: str) -> List[Dict]:
    """Extract ALL sublimits from Endt. No. 5(ii) section dynamically - handles multiple formats"""
    sublimits = []
    
    # Look for Endt. No. 5(ii) patterns
    endorsement_5ii_match = re.search(r'Endt\.\s*No\.\s*5\(ii\).*?(?=Endt\.\s*No\.|Endorsement\s*No\.|$)', text, re.IGNORECASE | re.DOTALL)
    if endorsement_5ii_match:
        endorsement_5ii_text = endorsement_5ii_match.group(0)
        print(f"[DEBUG] Found Endt. No. 5(ii) section")
        
        # SMART DETECTION: Check input format to choose extraction method
        print(f"[DEBUG] Analyzing input format...")
        
        # Check if this is the MERGED format (comma-separated conditions in one row)
        # Look for the specific merged string pattern with commas
        merged_conditions_pattern = r'Balloon\s*Sinsuplasty,.*?bronchical\s+thermoplasty,.*?vaporization.*?intra\s+operative\s+neuro\s+monitoring,.*?Intra\s+vitreal\s+injections'
        merged_format_detected = bool(re.search(merged_conditions_pattern, endorsement_5ii_text, re.IGNORECASE | re.DOTALL))
        
        print(f"[DEBUG] Merged format detection result: {merged_format_detected}")
        if merged_format_detected:
            print(f"[DEBUG] Found merged pattern in text")
        
        # Additional check: if we find specific merged condition patterns, force merged format
        force_merged = bool(re.search(
            r'(?:Treatment\s+of\s+mental\s+illness.*?neurodegenerative|Oral\s+Chemotherapy.*?Immunotherapy.*?monoclonal)',
            endorsement_5ii_text, re.IGNORECASE | re.DOTALL
        ))
        
        if merged_format_detected or force_merged:
            print(f"[DEBUG] MERGED FORMAT detected - using merged extraction logic")
            return extract_merged_format(endorsement_5ii_text)
        else:
            print(f"[DEBUG] INDIVIDUAL FORMAT detected - using individual extraction logic")
            return extract_individual_format(endorsement_5ii_text)
    
    return sublimits



def extract_merged_format(endorsement_5ii_text: str) -> List[Dict]:
    """Extract sublimits for MERGED format (comma-separated conditions stay merged)"""
    sublimits = []
    
    # Extract conditions in the correct order based on input sequence
    # Define all conditions with their patterns in the order they appear in input
    all_conditions = [
        {
            "name": "Cataract",
            "patterns": [
                r'\bCataract\b',
                r'Treatment\s*\n\s*Cataract',
                r'Treatment\s+Cataract',
                r'Cataract\s+Nil\s+Capping',
                r'Cataract.*?Nil\s+Capping',
                r'Cataract\s*\|.*?Nil\s+Capping',
                r'Cataract\s*\|.*?Nil'
            ]
        },
        {
            "name": "Treatment of mental illness, stress or psychological disorders and neurodegenerative disorders",
            "patterns": [
                r'Treatment\s+of\s+mental\s+illness,?\s+stress\s+or\s+psychological\s+disorders\s+and\s+neurodegenerative\s+disorders',
                r'Treatment\s+of\s+mental\s+illness,\s*stress\s*or\s*psychological\s*disorders\s*and\s*neurodegenerative\s*disorders',
                r'Treatment\s+of\s+mental\s+illness.*?neurodegenerative\s+disorders',
                r'Treatment\s+of\s+mental\s+illness.*?psychological\s+disorders.*?neurodegenerative\s+disorders',
                r'mental\s+illness.*?psychological\s+disorders.*?neurodegenerative\s+disorders',
                r'Treatment\s+of\s+mental\s+illness.*?stress.*?psychological.*?neurodegenerative',
                r'mental\s+illness.*?stress.*?psychological.*?neurodegenerative'
            ]
        },
        {
            "name": "Balloon Sinsuplasty, bronchical thermoplasty, vaporization of prostate(green laser treatment), intra operative neuro monitoring, Intra vitreal injections",
            "patterns": [
                r'Balloon\s*Sinsuplasty,\s*bronchical\s+thermoplasty,\s*vaporization\s+of\s+prostate\s*\([^)]*\),?\s*intra\s+operative\s+neuro\s+monitoring,?\s*Intra\s+vitreal\s+injections',
                r'Balloon\s*Sinsuplasty,\s*bronchical\s+thermoplasty,\s*vaporization.*?intra\s+operative\s+neuro\s+monitoring.*?Intra\s+vitreal\s+injections',
                r'Balloon.*?thermoplasty.*?vaporization.*?monitoring.*?injections',
                r'Balloon\s*Sinsuplasty,.*?bronchical\s+thermoplasty,.*?vaporization.*?intra\s+operative\s+neuro\s+monitoring.*?Intra\s+vitreal\s+injections'
            ]
        },
        {
            "name": "Stem Cell therapy",
            "patterns": [
                r'Stem\s+Cell\s+therapy\s*-\s*Hematopoietic\s+stem\s+cells\s+for\s+bone\s+marrow\s+transplant\s+for\s+haematological\s+conditions\s+to\s+be\s+covered',
                r'Stem\s+Cell\s+therapy\s*-\s*Hematopoietic\s+stem\s+cells\s+for\s+bone\s+marrow\s+transplant\s+for\s+h[ae]ematological\s+conditions\s+to\s+be\s+covered',
                r'Stem\s+Cell\s+therapy.*?Hematopoietic.*?bone\s+marrow\s+transplant.*?haematological\s+conditions.*?covered',
                r'Stem\s+Cell\s+therapy.*?bone\s+marrow\s+transplant.*?covered',
                r'Stem\s+Cell\s+therapy\s*-\s*Hematopoietic.*?bone.*?marrow.*?transplant.*?haematological.*?conditions.*?covered',
                r'Stem\s+Cell\s+therapy.*?Hematopoietic.*?transplant.*?conditions.*?covered',
                r'Stem\s+Cell\s+therapy.*?bone\s+marrow.*?transplant.*?covered',
                r'Stem\s+Cell\s+therapy.*?haematological.*?conditions.*?covered',
                r'Stem\s+Cell\s+therapy'
            ]
        },
        {
            "name": "Oral Chemotherapy, Immunotherapy(monoclonal antibody to be given as injection)",
            "patterns": [
                r'Oral\s+Chemotherapy,?\s+Immunotherapy\s*\(\s*monoclonal\s+antibody\s+to\s+be\s+given\s+as\s+injection\s*\)',
                r'Oral\s+Chemotherapy,?\s+Immunotherapy\s*\([^)]*monoclonal[^)]*\)',
                r'Oral\s+Chemotherapy,?\s+Immunotherapy\s*\([^)]*\)',
                r'Oral\s+Chemotherapy,?\s+Immunotherapy',
                r'Oral\s+Chemotherapy.*?Immunotherapy.*?monoclonal.*?antibody',
                r'Oral\s+Chemotherapy.*?Immunotherapy.*?injection',
                r'Oral\s+Chemotherapy.*?Immunotherapy'
            ]
        }
    ]
    
    # Extract conditions in order
    for condition in all_conditions:
        condition_name = condition["name"]
        patterns = condition["patterns"]
        
        # Special debug for Stem Cell therapy
        if "Stem Cell" in condition_name:
            print(f"[DEBUG] Searching for Stem Cell therapy...")
            print(f"[DEBUG] Will test {len(patterns)} patterns")
        
        # Special debug for Cataract
        if condition_name == "Cataract":
            print(f"[DEBUG] Searching for Cataract...")
            print(f"[DEBUG] Will test {len(patterns)} patterns for Cataract")
            # Check if Cataract appears anywhere in the text
            if re.search(r'[Cc]ataract', endorsement_5ii_text, re.IGNORECASE):
                print(f"[DEBUG] Cataract found in text")
            else:
                print(f"[DEBUG] WARNING: Cataract NOT found in text at all")
        
        for pattern_idx, pattern in enumerate(patterns):
            if re.search(pattern, endorsement_5ii_text, re.IGNORECASE | re.DOTALL):
                print(f"[DEBUG] Found condition: {condition_name}")
                
                # Check if already exists to avoid duplicates
                if any(sub['type'] == condition_name for sub in sublimits):
                    print(f"[DEBUG] Condition {condition_name} already exists, skipping")
                    break
                
                # Try to find associated amount and percentage
                limit_amount = ""
                percentage = ""  # Will extract from text
                
                # ENHANCED: Check for "Nil Capping" specifically for Cataract
                if condition_name == "Cataract":
                    nil_capping_match = re.search(r'Cataract.*?Nil\s+Capping', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                    if nil_capping_match:
                        print(f"[DEBUG] Found Cataract with Nil Capping - setting limit to empty")
                        limit_amount = ""
                        sublimits.append({
                            "applicable": "Yes",
                            "type": condition_name,
                            "limit": limit_amount
                        })
                        break
                    
                    # CRITICAL FIX: Check for Cataract with missing percentage (just "%" without number)
                    missing_percentage_match = re.search(r'Cataract.*?%\s+of\s+the\s+sum\s+insured.*?maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)/-', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                    if missing_percentage_match:
                        print(f"[DEBUG] Found Cataract with missing percentage (just '%' without number) - setting limit to empty")
                        limit_amount = ""
                        sublimits.append({
                            "applicable": "Yes",
                            "type": condition_name,
                            "limit": limit_amount
                        })
                        break
                
                # Look for specific amount patterns near the condition
                # First try to find the condition in context and extract amount from nearby text
                condition_context = re.search(f'{pattern}.*?(?:\n|$)', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                if condition_context:
                    context_text = condition_context.group(0)
                    
                    # ENHANCED: Check for "Nil Capping" in the context text first
                    if condition_name == "Cataract" and re.search(r'Nil\s+Capping', context_text, re.IGNORECASE):
                        print(f"[DEBUG] Found Cataract with Nil Capping in context - setting limit to empty")
                        limit_amount = ""
                    else:
                        # Look for percentage first
                        percentage_match = re.search(r'(\d+)%\s+of\s+the\s+sum\s+insured', context_text, re.IGNORECASE)
                        if percentage_match:
                            percentage = f"{percentage_match.group(1)}%"
                            print(f"[DEBUG] Found percentage: {percentage}")
                        
                        # Look for specific amount patterns in the context
                        amount_patterns_in_context = [
                            r'sublimit\s+of\s+(\d+)',  # "sublimit of 30000"
                            r'(?:INR\.?|Rs\.?)\s*([\d,]+)/-',  # "Rs.100000/-"
                            r'maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)',  # "maximum of Rs.100000"
                            r'(\d+)\s*during\s+the\s+Period\s+of\s+Insurance',  # "100000 during the Period of Insurance"
                            r'(?:INR\.?|Rs\.?)\s*([\d,]+)',  # General amount pattern
                        ]
                        
                        for amount_pattern in amount_patterns_in_context:
                            amount_match = re.search(amount_pattern, context_text, re.IGNORECASE)
                            if amount_match:
                                limit_amount = amount_match.group(1).replace(',', '')
                                print(f"[DEBUG] Found amount in context: {limit_amount}")
                                break
                
                # If no amount found in context, try to find it in the broader text
                if not limit_amount:
                    # ENHANCED: Check for "Nil Capping" in broader text for Cataract
                    if condition_name == "Cataract":
                        broader_nil_match = re.search(r'Cataract.*?Nil\s+Capping', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                        if broader_nil_match:
                            print(f"[DEBUG] Found Cataract with Nil Capping in broader text - setting limit to empty")
                            limit_amount = ""
                        else:
                            # Look for percentage and amount patterns near the condition
                            context_match = re.search(f'{pattern}.*?(?:INR\.?|Rs\.?)\s*([\d,]+)', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                            if context_match:
                                limit_amount = context_match.group(1).replace(',', '')
                                print(f"[DEBUG] Found amount: {limit_amount}")
                            
                            # CRITICAL FIX: Only look for percentage in the immediate context of Cataract, not broader text
                            # This prevents picking up percentages from other conditions
                            cataract_context = re.search(r'Cataract.*?(?:\n|$)', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                            if cataract_context:
                                cataract_text = cataract_context.group(0)
                                percentage_context = re.search(r'(\d+)%\s+of\s+the\s+sum\s+insured', cataract_text, re.IGNORECASE)
                                if percentage_context:
                                    percentage = f"{percentage_context.group(1)}%"
                                    print(f"[DEBUG] Found percentage in Cataract context: {percentage}")
                                else:
                                    print(f"[DEBUG] No percentage found in Cataract context - setting to empty")
                                    percentage = ""
                            else:
                                print(f"[DEBUG] No Cataract context found - setting percentage to empty")
                                percentage = ""
                    else:
                        # Look for percentage and amount patterns near the condition
                        context_match = re.search(f'{pattern}.*?(?:INR\.?|Rs\.?)\s*([\d,]+)', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                        if context_match:
                            limit_amount = context_match.group(1).replace(',', '')
                            print(f"[DEBUG] Found amount: {limit_amount}")
                        
                        # CRITICAL FIX: Only look for percentage in the immediate context of the condition, not broader text
                        # This prevents picking up percentages from other conditions
                        condition_context = re.search(f'{pattern}.*?(?:\n|$)', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                        if condition_context:
                            condition_text = condition_context.group(0)
                            percentage_context = re.search(r'(\d+)%\s+of\s+the\s+sum\s+insured', condition_text, re.IGNORECASE)
                            if percentage_context:
                                percentage = f"{percentage_context.group(1)}%"
                                print(f"[DEBUG] Found percentage in {condition_name} context: {percentage}")
                            else:
                                print(f"[DEBUG] No percentage found in {condition_name} context - keeping existing")
                                # Keep the existing percentage if already found in earlier context
                        else:
                            print(f"[DEBUG] No {condition_name} context found - keeping existing percentage")
                            # Keep the existing percentage if already found in earlier context
                
                # If we found a percentage, calculate the amount using it
                if percentage and limit_amount:
                    try:
                        # Extract the maximum amount from the limit_amount
                        max_amount_match = re.search(r'(\d+)', limit_amount)
                        if max_amount_match:
                            max_amount = int(max_amount_match.group(1))
                            # Calculate using the percentage and sum insured
                            percent_value = float(percentage.strip('%')) / 100
                            calculated_amount = int(500000 * percent_value)
                            # Return the smaller value between calculated and maximum
                            final_amount = min(calculated_amount, max_amount)
                            limit_amount = str(final_amount)
                            print(f"[DEBUG] Calculated amount using {percentage}: {calculated_amount}, max: {max_amount}, final: {limit_amount}")
                    except:
                        pass  # Keep original amount if calculation fails
                elif not percentage and limit_amount:
                    # If no percentage found but amount exists, use the amount as is
                    print(f"[DEBUG] No percentage found for {condition_name}, using amount as is: {limit_amount}")
                elif percentage and not limit_amount:
                    # If percentage found but no amount, calculate using percentage only
                    try:
                        percent_value = float(percentage.strip('%')) / 100
                        calculated_amount = int(500000 * percent_value)
                        limit_amount = str(calculated_amount)
                        print(f"[DEBUG] Calculated amount using {percentage} only: {limit_amount}")
                    except:
                        limit_amount = ""
                        print(f"[DEBUG] Failed to calculate amount using {percentage}")
                else:
                    # No percentage and no amount - keep empty
                    print(f"[DEBUG] No percentage and no amount found for {condition_name} - keeping empty")
                    limit_amount = ""
                
                sublimits.append({
                    "applicable": "Yes",
                    "type": condition_name,
                    "limit": limit_amount
                })
                break
            else:
                # Special debug for Stem Cell therapy
                if "Stem Cell" in condition_name:
                    print(f"[DEBUG] Pattern {pattern_idx + 1} did not match for Stem Cell therapy")
        
        # Special debug for Stem Cell therapy if no patterns matched
        if "Stem Cell" in condition_name and not any(sub['type'] == condition_name for sub in sublimits):
            print(f"[DEBUG] No Stem Cell therapy patterns matched!")
            # Check if 'Stem Cell' appears anywhere in the text
            if re.search(r'Stem\s+Cell', endorsement_5ii_text, re.IGNORECASE):
                print(f"[DEBUG] 'Stem Cell' found in text - checking surrounding context")
                stem_match = re.search(r'Stem\s+Cell.*?(?:\n.*?){0,5}', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                if stem_match:
                    print(f"[DEBUG] Stem Cell context: {stem_match.group(0)}")
            else:
                print(f"[DEBUG] 'Stem Cell' NOT found in text at all")
    
    # Method 3: Extract amounts from general amount patterns if not found yet
    # Look for specific amount patterns in the text
    amount_patterns = [
        r'sublimit\s+of\s+(\d+)',  # "sublimit of 30000"
        r'(?:INR\.?|Rs\.?)\s*([\d,]+)/-',  # "Rs.100000/-"
        r'maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)',  # "maximum of Rs.100000"
        r'(?:INR\.?|Rs\.?)\s*([\d,]+)',  # General amount pattern
        r'(\d+)\s*during\s+the\s+Period\s+of\s+Insurance'  # "100000 during the Period of Insurance"
    ]
    
    # Try to extract amounts for each sublimit that doesn't have one yet
    for sublimit in sublimits:
        if not sublimit["limit"]:  # If no amount found yet
            # ENHANCED: Check for "Nil Capping" for Cataract before assigning amounts
            if sublimit["type"] == "Cataract":
                nil_capping_check = re.search(r'Cataract.*?Nil\s+Capping', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                if nil_capping_check:
                    print(f"[DEBUG] Cataract has Nil Capping - keeping limit empty")
                    continue  # Skip amount assignment for Cataract with Nil Capping
                
                # CRITICAL FIX: Also check for Cataract with missing percentage (just "%" without number)
                missing_percentage_check = re.search(r'Cataract.*?%\s+of\s+the\s+sum\s+insured.*?maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)/-', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                if missing_percentage_check:
                    print(f"[DEBUG] Cataract has missing percentage - keeping limit empty")
                    continue  # Skip amount assignment for Cataract with missing percentage
            
            # Try to find amount in the broader text context
            for amount_pattern in amount_patterns:
                amount_match = re.search(amount_pattern, endorsement_5ii_text, re.IGNORECASE)
                if amount_match:
                    sublimit["limit"] = amount_match.group(1).replace(',', '')
                    print(f"[DEBUG] Assigned amount {sublimit['limit']} to {sublimit['type']}")
                    break
    
    # Post-process to assign correct amounts based on condition types
    # This ensures we get the right amounts for each condition based on the input format
    for sublimit in sublimits:
        condition_type = sublimit["type"].lower()
        
        # CRITICAL FIX: Check for Cataract with Nil Capping first before any amount assignment
        if condition_type == "cataract" or "cataract" in condition_type:
            nil_capping_check = re.search(r'Cataract.*?Nil\s+Capping', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
            if nil_capping_check:
                print(f"[DEBUG] Cataract has Nil Capping - keeping limit empty in post-processing")
                sublimit["limit"] = ""  # Ensure it stays empty
                continue  # Skip all amount assignment for Cataract with Nil Capping
            
            # CRITICAL FIX: Also check for Cataract with missing percentage (just "%" without number)
            missing_percentage_check = re.search(r'Cataract.*?%\s+of\s+the\s+sum\s+insured.*?maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)/-', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
            if missing_percentage_check:
                print(f"[DEBUG] Cataract has missing percentage - keeping limit empty in post-processing")
                sublimit["limit"] = ""  # Ensure it stays empty
                continue  # Skip all amount assignment for Cataract with missing percentage
        
        # CRITICAL FIX: Also check if Cataract already has an amount but should be empty due to Nil Capping or missing percentage
        if condition_type == "cataract" or "cataract" in condition_type:
            nil_capping_check = re.search(r'Cataract.*?Nil\s+Capping', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
            missing_percentage_check = re.search(r'Cataract.*?%\s+of\s+the\s+sum\s+insured.*?maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)/-', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
            if (nil_capping_check or missing_percentage_check) and sublimit["limit"]:  # If it has a limit but should be empty
                print(f"[DEBUG] Cataract has Nil Capping or missing percentage but has amount {sublimit['limit']} - clearing it")
                sublimit["limit"] = ""  # Clear any existing amount
                continue  # Skip all amount assignment for Cataract with Nil Capping or missing percentage
        
        # Assign amounts based on condition type if not already set
        if not sublimit["limit"]:
            # First, try to find any percentage pattern for this condition
            condition_name_for_search = condition_type.replace(' ', '.*?')
            percentage_pattern = rf'{condition_name_for_search}.*?(\d+)%\s+of\s+the\s+sum\s+insured'
            percentage_match = re.search(percentage_pattern, endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
            
            if percentage_match:
                percentage = f"{percentage_match.group(1)}%"
                # Look for maximum amount
                max_amount_match = re.search(r'maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)/-', endorsement_5ii_text, re.IGNORECASE)
                if max_amount_match:
                    max_amount = int(max_amount_match.group(1).replace(',', ''))
                    sublimit["limit"] = str(max_amount)
                    print(f"[DEBUG] Using maximum amount for {condition_type}: {sublimit['limit']}")
                else:
                    # If no maximum found, calculate using the percentage and sum insured
                    if percentage:
                        try:
                            percent_value = float(percentage.strip('%')) / 100
                            calculated_amount = int(500000 * percent_value)
                            sublimit["limit"] = str(calculated_amount)
                            print(f"[DEBUG] Calculated amount using {percentage} and sum insured 500000: {sublimit['limit']}")
                        except:
                            sublimit["limit"] = ""
                            print(f"[DEBUG] Failed to calculate amount for {condition_type}")
                    else:
                        sublimit["limit"] = ""
                        print(f"[DEBUG] No percentage found for {condition_type}")
            else:
                # Fallback to specific condition patterns
                if "mental illness" in condition_type or "psychological disorders" in condition_type:
                    # Fallback to "sublimit of 30000" pattern
                    mental_match = re.search(r'sublimit\s+of\s+(\d+)', endorsement_5ii_text, re.IGNORECASE)
                    if mental_match:
                        sublimit["limit"] = mental_match.group(1)
                        print(f"[DEBUG] Assigned mental illness amount: {sublimit['limit']}")
                
                elif "balloon" in condition_type and "thermoplasty" in condition_type and "vaporization" in condition_type and "monitoring" in condition_type and "injections" in condition_type:
                    # Fallback to general maximum pattern
                    balloon_match = re.search(r'maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)/-', endorsement_5ii_text, re.IGNORECASE)
                    if balloon_match:
                        sublimit["limit"] = balloon_match.group(1).replace(',', '')
                        print(f"[DEBUG] Assigned balloon/thermoplasty/vaporization/monitoring/injections amount: {sublimit['limit']}")
                
                elif "stem cell" in condition_type:
                    # Fallback to general maximum pattern
                    stem_match = re.search(r'maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)/-', endorsement_5ii_text, re.IGNORECASE)
                    if stem_match:
                        sublimit["limit"] = stem_match.group(1).replace(',', '')
                        print(f"[DEBUG] Assigned stem cell amount: {sublimit['limit']}")
                
                elif "oral chemotherapy" in condition_type or "immunotherapy" in condition_type:
                    # Fallback to "during the Period of Insurance" pattern
                    chemo_match = re.search(r'(\d+)\s*during\s+the\s+Period\s+of\s+Insurance', endorsement_5ii_text, re.IGNORECASE)
                    if chemo_match:
                        sublimit["limit"] = chemo_match.group(1)
                        print(f"[DEBUG] Assigned chemotherapy amount: {sublimit['limit']}")
    
    # Final pass: Ensure all sublimits have the correct amounts
    # This handles cases where the amounts might be in different formats
    for sublimit in sublimits:
        if not sublimit["limit"]:
            # CRITICAL FIX: Check for "Nil Capping" for Cataract before fallback assignment
            if sublimit["type"] == "Cataract" or "cataract" in sublimit["type"].lower():
                nil_capping_final_check = re.search(r'Cataract.*?Nil\s+Capping', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                if nil_capping_final_check:
                    print(f"[DEBUG] Cataract has Nil Capping - keeping limit empty in final pass")
                    sublimit["limit"] = ""  # Ensure it stays empty
                    continue  # Skip fallback amount assignment for Cataract with Nil Capping
                
                # CRITICAL FIX: Also check for Cataract with missing percentage (just "%" without number)
                missing_percentage_final_check = re.search(r'Cataract.*?%\s+of\s+the\s+sum\s+insured.*?maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)/-', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                if missing_percentage_final_check:
                    print(f"[DEBUG] Cataract has missing percentage - keeping limit empty in final pass")
                    sublimit["limit"] = ""  # Ensure it stays empty
                    continue  # Skip fallback amount assignment for Cataract with missing percentage
            
            # Try to find any amount in the text and assign it
            all_amounts = re.findall(r'(?:INR\.?|Rs\.?)\s*([\d,]+)', endorsement_5ii_text, re.IGNORECASE)
            if all_amounts:
                # Use the first available amount
                sublimit["limit"] = all_amounts[0].replace(',', '')
                print(f"[DEBUG] Assigned fallback amount {sublimit['limit']} to {sublimit['type']}")

    print(f"[DEBUG] Total sublimits extracted: {len(sublimits)}")
    for i, sublimit in enumerate(sublimits):
        print(f"[DEBUG] {i+1}. Type: '{sublimit['type']}', Limit: '{sublimit['limit']}'")
    
    # Debug: Check what conditions we missed
    expected_conditions = [
        "Cataract",
        "Treatment of mental illness, stress or psychological disorders and neurodegenerative disorders",
        "Balloon Sinsuplasty, bronchical thermoplasty, vaporization of prostate(green laser treatment), intra operative neuro monitoring, Intra vitreal injections",
        "Stem Cell therapy",
        "Oral Chemotherapy, Immunotherapy(monoclonal antibody to be given as injection)"
    ]
    
    found_conditions = [sub['type'] for sub in sublimits]
    missing = [cond for cond in expected_conditions if cond not in found_conditions]
    if missing:
        print(f"[DEBUG] Missing conditions: {missing}")
        
        # Try to find missing conditions with more flexible patterns
        for missing_condition in missing:
            if "mental illness" in missing_condition.lower():
                # Look for any mention of mental illness
                if re.search(r'mental\s+illness', endorsement_5ii_text, re.IGNORECASE):
                    print(f"[DEBUG] Found mental illness mention, adding condition")
                    sublimits.append({
                        "applicable": "Yes",
                        "type": missing_condition,
                        "limit": ""
                    })
            
            elif "oral chemotherapy" in missing_condition.lower() or "immunotherapy" in missing_condition.lower():
                # Look for any mention of chemotherapy or immunotherapy
                if re.search(r'(?:oral\s+)?chemotherapy|immunotherapy', endorsement_5ii_text, re.IGNORECASE):
                    print(f"[DEBUG] Found chemotherapy/immunotherapy mention, adding condition")
                    sublimits.append({
                        "applicable": "Yes",
                        "type": missing_condition,
                        "limit": ""
                    })
            
            elif "stem cell" in missing_condition.lower():
                # Look for any mention of stem cell
                if re.search(r'stem\s+cell', endorsement_5ii_text, re.IGNORECASE):
                    print(f"[DEBUG] Found stem cell mention, adding condition")
                    sublimits.append({
                        "applicable": "Yes",
                        "type": missing_condition,
                        "limit": ""
                    })
            
            elif "balloon" in missing_condition.lower() and "thermoplasty" in missing_condition.lower() and "vaporization" in missing_condition.lower() and "monitoring" in missing_condition.lower() and "injections" in missing_condition.lower():
                # Look for any mention of the merged condition
                if re.search(r'balloon.*?thermoplasty.*?vaporization.*?monitoring.*?injections', endorsement_5ii_text, re.IGNORECASE):
                    print(f"[DEBUG] Found balloon/thermoplasty/vaporization/monitoring/injections mention, adding condition")
                    sublimits.append({
                        "applicable": "Yes",
                        "type": missing_condition,
                        "limit": ""
                    })
    
    # Post-processing: Split merged conditions into individual ones for better output format
    # This preserves the original logic while providing the desired output format
    final_sublimits = []
    
    for sublimit in sublimits:
        condition_name = sublimit["type"]
        
        # Check if this is the mental illness merged condition that should be split
        if "Treatment of mental illness, stress or psychological disorders and neurodegenerative disorders" in condition_name:
            # Split into individual conditions
            individual_conditions = [
                "Treatment of mental illness",
                "stress or psychological disorders",
                "neurodegenerative disorders"
            ]
            
            # Use the original calculated amount from the merged condition
            original_amount = sublimit["limit"]
            print(f"[DEBUG] Using original amount from merged condition: {original_amount}")
            
            # Create individual sublimit entries for each condition
            for individual_condition in individual_conditions:
                # Use the original amount for all mental illness related conditions
                individual_amount = original_amount
                
                # Special handling for mental illness conditions to ensure correct amount
                if "mental illness" in individual_condition.lower() or "psychological disorders" in individual_condition.lower() or "neurodegenerative disorders" in individual_condition.lower():
                    # Look for the specific "sublimit of 30000" pattern
                    mental_match = re.search(r'sublimit\s+of\s+(\d+)', endorsement_5ii_text, re.IGNORECASE)
                    if mental_match:
                        individual_amount = mental_match.group(1)
                        print(f"[DEBUG] Corrected mental illness amount for {individual_condition}: {individual_amount}")
                
                final_sublimits.append({
                    "applicable": "Yes",
                    "type": individual_condition,
                    "limit": individual_amount
                })
        
        # Check if this is the Oral Chemotherapy merged condition that should be split
        elif "Oral Chemotherapy, Immunotherapy(monoclonal antibody to be given as injection)" in condition_name:
            # Split into individual conditions
            individual_conditions = [
                "Oral Chemotherapy",
                "Immunotherapy"
            ]
            
            # Use the original calculated amount from the merged condition
            original_amount = sublimit["limit"]
            print(f"[DEBUG] Using original amount from merged condition: {original_amount}")
            
            # Create individual sublimit entries for each condition
            for individual_condition in individual_conditions:
                # Use the original amount for all chemotherapy/immunotherapy related conditions
                individual_amount = original_amount
                
                # Special handling for Oral Chemotherapy and Immunotherapy conditions
                if any(keyword in individual_condition.lower() for keyword in ["oral chemotherapy", "immunotherapy"]):
                    # Look for the percentage pattern for the merged condition (5% per month)
                    percentage_match = re.search(r'(\d+)%\s+of\s+the\s+Sum\s+Insured\s+per\s+month\s+subject\s+to\s+a\s+maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)/-', endorsement_5ii_text, re.IGNORECASE)
                    if percentage_match:
                        percentage = int(percentage_match.group(1))
                        max_amount = int(percentage_match.group(2).replace(',', ''))
                        # Calculate using the percentage and sum insured
                        calculated_amount = int((percentage / 100) * 500000)
                        # Return the smaller value between calculated and maximum
                        final_amount = min(calculated_amount, max_amount)
                        individual_amount = str(final_amount)
                        print(f"[DEBUG] Corrected Oral Chemotherapy/Immunotherapy amount for {individual_condition}: {individual_amount} (using {percentage}%)")
                
                final_sublimits.append({
                    "applicable": "Yes",
                    "type": individual_condition,
                    "limit": individual_amount
                })
        
        # Check if this is the balloon/thermoplasty merged condition that should be split
        elif "Balloon Sinsuplasty, bronchical thermoplasty, vaporization of prostate(green laser treatment), intra operative neuro monitoring, Intra vitreal injections" in condition_name:
            # Split into individual conditions
            individual_conditions = [
                "Balloon Sinsuplasty",
                "Bronchial Thermoplasty", 
                "Vaporization of prostate",
                "Intra Operative Neuro Monitoring",
                "Intra vitreal injections"
            ]
            
            # Use the original calculated amount from the merged condition
            original_amount = sublimit["limit"]
            print(f"[DEBUG] Using original amount from merged condition: {original_amount}")
            
            # Create individual sublimit entries for each condition
            for individual_condition in individual_conditions:
                # Use the original amount for all balloon/thermoplasty/vaporization/monitoring/injections related conditions
                individual_amount = original_amount
                
                # Special handling for balloon/thermoplasty/vaporization/monitoring/injections conditions
                if any(keyword in individual_condition.lower() for keyword in ["balloon", "thermoplasty", "vaporization", "monitoring", "injections", "bronchial", "intra operative", "intra vitreal"]):
                    # Look for the percentage pattern for the merged condition
                    percentage_match = re.search(r'(\d+)%\s+of\s+the\s+Sum\s+Insured\s+subject\s+to\s+a\s+maximum\s+of\s+(?:INR\.?|Rs\.?)\s*([\d,]+)/-', endorsement_5ii_text, re.IGNORECASE)
                    if percentage_match:
                        percentage = int(percentage_match.group(1))
                        max_amount = int(percentage_match.group(2).replace(',', ''))
                        # Calculate using the percentage and sum insured
                        calculated_amount = int((percentage / 100) * 500000)
                        # Return the smaller value between calculated and maximum
                        final_amount = min(calculated_amount, max_amount)
                        individual_amount = str(final_amount)
                        print(f"[DEBUG] Corrected balloon/thermoplasty/vaporization/monitoring/injections amount for {individual_condition}: {individual_amount} (using {percentage}%)")
                
                final_sublimits.append({
                    "applicable": "Yes",
                    "type": individual_condition,
                    "limit": individual_amount
                })
        else:
            # For non-merged conditions, clean up the name and keep as is
            cleaned_condition_name = condition_name
            # Remove brackets and their contents
            cleaned_condition_name = re.sub(r'\s*\([^)]*\)', '', cleaned_condition_name)
            # Remove hyphens
            cleaned_condition_name = re.sub(r'\s*-\s*', ' ', cleaned_condition_name)
            # Clean up extra spaces
            cleaned_condition_name = re.sub(r'\s+', ' ', cleaned_condition_name).strip()
            
            final_sublimits.append({
                "applicable": sublimit["applicable"],
                "type": cleaned_condition_name,
                "limit": sublimit["limit"]
            })
    
    return final_sublimits





def extract_individual_format(endorsement_5ii_text: str) -> List[Dict]:
    """Extract sublimits for INDIVIDUAL format (each condition gets separate row)"""
    print(f"[DEBUG] Starting INDIVIDUAL format extraction")
    sublimits = []
    table_row_patterns = [
        # Pattern 1: Exact condition names from table (case-insensitive) - ENHANCED PATTERNS
        r'^\s*(Cataract)\s*(?:\||Nil|50%|5\(ii\)|$)',
        r'^\s*(Cataract)\s*Nil\s*Capping',
        r'^\s*(Cataract)\s*\|.*?Nil\s*Capping',
        r'^\s*(Cataract)\s*\|.*?Nil',
        r'^\s*(Cataract)\s*Nil',
        r'^\s*(Balloon\s+Sinsuplasty)\s*(?:\||50%|5\(ii\)|$)',
        r'^\s*(Stem\s+Cell\s+therapy)\s*(?:\||50%|5\(ii\)|$)',
        r'^\s*(Oral\s+Chemotherapy,?\s+Immunotherapy\s*\(\s*monoclonal\s+antibody\s+to\s+be\s+given\s+as\s+injection\s*\))\s*(?:\||50%|5\(ii\)|$)',
        r'^\s*(Bronchical\s+Thermoplasty)\s*(?:\||50%|5\(ii\)|$)',
        
        # VAPORIZATION PATTERNS - individual line formats
        r'^\s*(vaporization\s+of\s+prostate\s*\(\s*green\s+laser\s+treatment\s*\))\s*(?:\||Nil|50%|10%|5\(ii\)|$)',
        r'^\s*(vaporization\s+of\s+prostate\s*\(\s*green\s+laser\s+treatment\s*\))',
        r'^\s*(vaporization\s+of\s+prostate\s*\([^)]*\))',
        r'^\s*(vaporization\s+of\s+prostate)',
        
        r'^\s*(Intra\s+Operative\s+Neuro\s+Monitoring)\s*(?:\||50%|5\(ii\)|$)',
        r'^\s*(Intra\s+vitreal\s+injections)\s*(?:\||50%|5\(ii\)|$)',
        
        # Individual mental illness conditions
        r'^\s*(Treatment\s+of\s+mental\s+illness)\s*(?:\||50%|5\(ii\)|$)',
        r'^\s*(stress\s+or\s+psychological\s+disorders)\s*(?:\||50%|5\(ii\)|$)',
        r'^\s*(neurodegenerative\s+disorders)\s*(?:\||50%|5\(ii\)|$)',
        
        # Individual chemotherapy conditions
        r'^\s*(Oral\s+Chemotherapy)\s*(?:\||50%|5\(ii\)|$)',
        r'^\s*(Immunotherapy)\s*(?:\||50%|5\(ii\)|$)',
        
        # Pattern 2: More flexible patterns for variations
        r'(Treatment\s+of\s+mental\s+illness,\s+stress\s+or\s+psychological\s+disorders\s+and\s+neurodegenerative\s+disorders\.?)',
        r'(Treatment\s+of\s+mental\s+illness)',
        r'(stress\s+or\s+psychological\s+disorders)',
        r'(neurodegenerative\s+disorders)',
        r'(Stem\s+[Cc]ell\s+[Tt]herapy(?:\s*-[^|]*)?)',
        r'(Oral\s+[Cc]hemotherapy,?\s*[Ii]mmunotherapy(?:\s*\([^)]+\))?)',
        r'(Oral\s+[Cc]hemotherapy)',
        r'([Ii]mmunotherapy)',
        r'(Balloon\s+[Ss]in[us]*plasty)',
        r'(Bronchical?\s+[Tt]hermoplasty)',
        r'(Intra\s+[Oo]perative\s+[Nn]euro\s+[Mm]onitoring)',
        r'(Intra\s+[Vv]itreal\s+[Ii]njections)',
        
        # Pattern 3: Condition followed by "Nil Capping" - enhanced for vaporization
        r'^([A-Za-z][A-Za-z\s,\-()]+?)\s+Nil\s+Capping',
        r'(vaporization[^|]+?)\s+Nil\s+Capping',
        
        # Pattern 4: Condition followed by percentage and amount description - enhanced for vaporization
        r'^([A-Za-z][A-Za-z\s,\-()]+?)\s+(?:50%|10%|25%|5%)\s+of\s+the\s+sum\s+insured',
        r'(vaporization[^|]+?)\s+(?:50%|10%|25%|5%)\s+of\s+the\s+sum\s+insured',
        
        # Pattern 5: Specific format from user's text - condition followed by percentage and maximum
        r'^([A-Za-z][A-Za-z\s,\-()]+?)\s+(?:50%|10%|25%|5%)\s+of\s+the\s+sum\s+insured\s+subject\s+to\s+a\s+maximum',
        r'^([A-Za-z][A-Za-z\s,\-()]+?)\s+(?:50%|10%|25%|5%)\s+of\s+the\s+sum\s+insured\s+subject\s+to\s+a\s+maximum\s+of',
        r'^([A-Za-z][A-Za-z\s,\-()]+?)\s+(?:50%|10%|25%|5%)\s+of\s+the\s+sum\s+insured\s+subject\s+to\s+a\s+maximum\s+of\s+(?:INR\.?|Rs\.?)\s*[\d,]+/-',
        
        # Pattern 5a: Condition followed by 5(ii) - specific to user's format
        r'^([A-Za-z][A-Za-z\s,\-()]+?)\s+5\(ii\)',
        r'^([A-Za-z][A-Za-z\s,\-()]+?)\s+5\(ii\)\s*$',
        
        # Pattern 6: Medical conditions with specific keywords (broader match)
        r'([A-Za-z\s,\-()]*(?:[Tt]herapy|[Pp]lasty|[Mm]onitoring|[Ii]njections|[Cc]hemotherapy|[Ii]mmunotherapy|[Cc]ell|[Pp]rostate|[Nn]euro|[Vv]itreal|[Cc]ataract|[Tt]reatment)[A-Za-z\s,\-()]*)',
        
    ]
    
    # Split the text into lines and process each line
    lines = endorsement_5ii_text.split('\n')
    print(f"[DEBUG] Processing {len(lines)} lines for individual extraction")
    
    for line in lines:
        line = line.strip()
        if not line or len(line) < 5:  # Skip empty or very short lines
            continue
            
        print(f"[DEBUG] Processing line: {line[:100]}...")  # Debug first 100 chars
        
        # Special debug for vaporization
        if "vaporization" in line.lower():
            print(f"[DEBUG] Found vaporization line: {line}")
            print(f"[DEBUG] Vaporization line will be tested against {len(table_row_patterns)} patterns")
        
        # Special debug for Cataract
        if "cataract" in line.lower():
            print(f"[DEBUG] Found Cataract line: {line}")
            print(f"[DEBUG] Cataract line will be tested against {len(table_row_patterns)} patterns")
        
        # Try to extract condition and limit from each line
        for pattern_idx, pattern in enumerate(table_row_patterns):
            match = re.search(pattern, line, re.IGNORECASE | re.MULTILINE)
            if match:
                condition_name = match.group(1).strip()
                
                # Debug for vaporization matches
                if "vaporization" in condition_name.lower():
                    print(f"[DEBUG] VAPORIZATION MATCH found with pattern {pattern_idx}: '{condition_name}'")
                    print(f"[DEBUG] Pattern was: {pattern}")
                
                # Debug for Cataract matches
                if "cataract" in condition_name.lower():
                    print(f"[DEBUG] CATARACT MATCH found with pattern {pattern_idx}: '{condition_name}'")
                    print(f"[DEBUG] Pattern was: {pattern}")
                
                # Clean up condition name
                condition_name = re.sub(r'^[\d\.\s\|]+', '', condition_name)  # Remove leading numbers/pipes
                condition_name = re.sub(r'\s*\|\s*.*$', '', condition_name)  # Remove anything after pipe
                condition_name = condition_name.strip()
                
                # Debug after cleanup
                if "vaporization" in condition_name.lower():
                    print(f"[DEBUG] VAPORIZATION condition after cleanup: '{condition_name}'")
                
                # Skip if condition name is too short, looks like a header, or contains unwanted patterns
                skip_patterns = [
                    'sublimit', 'type', 'limit', 'condition', 'claim amount', 'payable', 'treatment',
                    'disease', 'illness', 'medical condition', 'injury', 'subject to', 'terms',
                    'exclusions', 'conditions', 'contained', 'policy', 'endorsed', 'declared',
                    'agreed', 'otherwise', 'following', 'HG00', 'Group Health', 'UIN:', 'IRDA'
                ]
                
                if (len(condition_name) < 3 or 
                    any(skip in condition_name.lower() for skip in skip_patterns) or
                    re.match(r'^[0-9%\s\-/]+$', condition_name)):  # Skip if only numbers/symbols
                    continue
                
                # Check if we already have this condition
                existing_condition = False
                for existing in sublimits:
                    if existing['type'].lower().replace(' ', '') == condition_name.lower().replace(' ', ''):
                        existing_condition = True
                        break
                
                if existing_condition:
                    continue
                
                sublimit_info = {
                    "applicable": "Yes",
                    "type": condition_name,
                    "limit": ""
                }
                
                # Check if this line has "Nil Capping" anywhere
                if re.search(r'Nil\s+Capping|nil\s+capping|NIL\s+CAPPING', line, re.IGNORECASE):
                    sublimit_info["limit"] = ""
                    print(f"[DEBUG] Found condition: {condition_name} with Nil Capping")
                else:
                    # ENHANCED: For Cataract, double-check if "Nil Capping" appears anywhere in the broader text
                    if condition_name.lower() == "cataract":
                        nil_capping_broader = re.search(r'Cataract.*?Nil\s+Capping', endorsement_5ii_text, re.IGNORECASE | re.DOTALL)
                        if nil_capping_broader:
                            sublimit_info["limit"] = ""
                            print(f"[DEBUG] Found Cataract with Nil Capping in broader text - setting limit to empty")
                        else:
                            # Try to extract amount from the same line first
                            amount_match = re.search(r'(?:INR\.?|Rs\.?)\s*([\d,]+)', line, re.IGNORECASE)
                            if amount_match:
                                amount_str = amount_match.group(1).replace(',', '')
                                sublimit_info["limit"] = amount_str
                                print(f"[DEBUG] Found condition: {condition_name} with amount on same line: {amount_str}")
                            else:
                                # Look for amount in the next few lines after this condition
                                found_amount = False
                                for next_line_idx in range(lines.index(line) + 1, min(lines.index(line) + 5, len(lines))):
                                    next_line = lines[next_line_idx].strip()
                                    
                                    # Look for INR/Rs amount patterns
                                    amount_match = re.search(r'(?:INR\.?|Rs\.?)\s*([\d,]+)/?-?', next_line, re.IGNORECASE)
                                    if amount_match:
                                        amount_str = amount_match.group(1).replace(',', '')
                                        sublimit_info["limit"] = amount_str
                                        print(f"[DEBUG] Found condition: {condition_name} with amount in next line: {amount_str}")
                                        found_amount = True
                                        break
                                    
                                    # Also look for percentage patterns that might indicate the amount
                                    percentage_match = re.search(r'(\d+)%\s+of\s+the\s+sum\s+insured', next_line, re.IGNORECASE)
                                    if percentage_match:
                                        percentage = percentage_match.group(1)
                                        print(f"[DEBUG] Found percentage: {percentage}% for condition: {condition_name}")
                                
                                if not found_amount:
                                    sublimit_info["limit"] = ""
                                    print(f"[DEBUG] Found condition: {condition_name} with no amount")
                    else:
                        # Try to extract amount from the same line first
                        amount_match = re.search(r'(?:INR\.?|Rs\.?)\s*([\d,]+)', line, re.IGNORECASE)
                        if amount_match:
                            amount_str = amount_match.group(1).replace(',', '')
                            sublimit_info["limit"] = amount_str
                            print(f"[DEBUG] Found condition: {condition_name} with amount on same line: {amount_str}")
                        else:
                            # Look for amount in the next few lines after this condition
                            found_amount = False
                            for next_line_idx in range(lines.index(line) + 1, min(lines.index(line) + 5, len(lines))):
                                next_line = lines[next_line_idx].strip()
                                
                                # Look for INR/Rs amount patterns
                                amount_match = re.search(r'(?:INR\.?|Rs\.?)\s*([\d,]+)/?-?', next_line, re.IGNORECASE)
                                if amount_match:
                                    amount_str = amount_match.group(1).replace(',', '')
                                    sublimit_info["limit"] = amount_str
                                    print(f"[DEBUG] Found condition: {condition_name} with amount in next line: {amount_str}")
                                    found_amount = True
                                    break
                                
                                # Also look for percentage patterns that might indicate the amount
                                percentage_match = re.search(r'(\d+)%\s+of\s+the\s+sum\s+insured', next_line, re.IGNORECASE)
                                if percentage_match:
                                    percentage = percentage_match.group(1)
                                    print(f"[DEBUG] Found percentage: {percentage}% for condition: {condition_name}")
                            
                            if not found_amount:
                                sublimit_info["limit"] = ""
                                print(f"[DEBUG] Found condition: {condition_name} with no amount")
                
                sublimits.append(sublimit_info)
                break  # Found a match for this line, move to next line
    
    # Post-process to clean up condition names and merge related entries
    cleaned_sublimits = []
    condition_map = {}
    
    for sublimit in sublimits:
        condition_name = sublimit["type"]
        
        # Clean up condition names based on common patterns
        cleaned_name = condition_name
        
        # Handle specific cleanup cases - standardize the condition names based on exact table format
        condition_lower = condition_name.lower()
        
        # Handle exact matches from the table first - ENHANCED FOR VAPORIZATION
        if "stem cell therapy" in condition_lower:
            cleaned_name = "Stem Cell therapy"
        elif "oral chemotherapy, immunotherapy(monoclonal" in condition_lower:
            cleaned_name = "Oral Chemotherapy, Immunotherapy(monoclonal antibody to be given as injection)"
        # ENHANCED VAPORIZATION HANDLING - Handle different variations
        elif "vaporization of prostate(green laser treatment)" in condition_lower:
            cleaned_name = "vaporization of prostate"
        elif "vaporization of prostate(green laser" in condition_lower:
            cleaned_name = "vaporization of prostate"
        elif "vaporization of prostate" in condition_lower and "green laser" in condition_lower:
            cleaned_name = "vaporization of prostate"
        elif "vaporization of prostate" in condition_lower and "green" not in condition_lower:
            cleaned_name = "vaporization of prostate"
        # Handle any other vaporization variations
        elif "vaporization" in condition_lower and "prostate" in condition_lower:
            cleaned_name = "vaporization of prostate"
            print(f"[DEBUG] VAPORIZATION condition standardized to: '{cleaned_name}'")
        elif "stem cell therapy" in condition_lower or "bone marrow transplant" in condition_lower:
            cleaned_name = "Stem Cell therapy"
        elif "oral chemotherapy" in condition_lower and "immunotherapy" in condition_lower:
            cleaned_name = "Oral Chemotherapy, Immunotherapy(monoclonal antibody to be given as injection)"
        elif "bronchial thermoplasty" in condition_lower or "bronchical thermoplasty" in condition_lower:
            cleaned_name = "Bronchial Thermoplasty"
        elif "intra operative neuro monitoring" in condition_lower:
            cleaned_name = "Intra Operative Neuro Monitoring"
        elif "balloon" in condition_lower and ("sinuplasty" in condition_lower or "sinsuplasty" in condition_lower):
            cleaned_name = "Balloon Sinsuplasty"
        elif "treatment of mental illness" in condition_lower:
            cleaned_name = "Treatment of mental illness"
        elif "stress or psychological disorders" in condition_lower:
            cleaned_name = "stress or psychological disorders"
        elif "neurodegenerative disorders" in condition_lower:
            cleaned_name = "neurodegenerative disorders"
        elif "oral chemotherapy" in condition_lower:
            cleaned_name = "Oral Chemotherapy"
        elif "immunotherapy" in condition_lower:
            cleaned_name = "Immunotherapy"
        elif "intra vitreal injections" in condition_lower:
            cleaned_name = "Intra vitreal injections"
        elif "cataract" in condition_lower:
            cleaned_name = "Cataract"
        
        # Remove trailing commas and clean up
        cleaned_name = re.sub(r',\s*$', '', cleaned_name)
        cleaned_name = cleaned_name.strip()
        
        # Remove brackets and their contents, and hyphens
        cleaned_name = re.sub(r'\s*\([^)]*\)', '', cleaned_name)
        cleaned_name = re.sub(r'\s*-\s*', ' ', cleaned_name)
        cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()
        
        # Skip if this is a duplicate or sub-part of existing condition
        if cleaned_name not in condition_map:
            condition_map[cleaned_name] = {
                "applicable": sublimit["applicable"],
                "type": cleaned_name,
                "limit": sublimit["limit"]
            }
        else:
            # If we already have this condition, keep the one with a limit if available
            if sublimit["limit"] and not condition_map[cleaned_name]["limit"]:
                condition_map[cleaned_name]["limit"] = sublimit["limit"]
    
    # Convert back to list
    sublimits = list(condition_map.values())
    
    # Post-process to merge related conditions that should be together
    merged_sublimits = []
    processed_conditions = set()
    
    for sublimit in sublimits:
        condition_name = sublimit["type"].lower()
        
        # Skip if already processed
        if condition_name in processed_conditions:
            continue
            
        # Check if this is part of a larger merged condition
        if "mental illness" in condition_name or "neurodegenerative" in condition_name:
            # For individual format, we want to keep these as separate conditions
            # So we just add them as individual conditions
            pass
                
        # For individual format, we don't merge these conditions - they should be separate
        # So we just add them as individual conditions
                
        elif "oral chemotherapy" in condition_name or "immunotherapy" in condition_name:
            # For individual format, we want to keep these as separate conditions
            # So we just add them as individual conditions
            pass
        
        # If not part of a merged condition, add as is
        if condition_name not in processed_conditions:
            merged_sublimits.append(sublimit)
            processed_conditions.add(condition_name)
    
    sublimits = merged_sublimits
    
    print(f"[DEBUG] Total sublimits extracted: {len(sublimits)}")
    for i, sublimit in enumerate(sublimits):
        print(f"[DEBUG] {i+1}. Type: '{sublimit['type']}', Limit: '{sublimit['limit']}'")
    
    return sublimits


def determine_member_type(min_age_years: int, max_age_years: int) -> str:
    """Determine member type based on age range"""
    # If the age range is primarily for children (max age <= 18)
    if max_age_years <= 18:
        return "Child"
    # If the age range is for adults (min age >= 18)
    elif min_age_years >= 18:
        return "Adult"
    # If the range spans both child and adult ages, determine based on the majority
    else:
        # Calculate the midpoint of the age range
        mid_age = (min_age_years + max_age_years) / 2
        if mid_age < 18:
            return "Child"
        else:
            return "Adult"

def extract_corporate_buffer_applicability(text: str) -> str:
    """
    Checks if Endt. No. 10 or Endorsement No. 10 exists in the text.
    Returns 'Yes' if applicable, else 'No'.
    """
    if re.search(r'(Endt\.|Endorsement)\s*No\.?\s*10\b', text, re.IGNORECASE):
        return "Yes"
    return "No"


def calculate_sublimit(sum_insured: int, percentage: str = "10%", max_amount: int = 500000) -> int:
    """
    Calculate sublimit based on sum insured, percentage string (e.g., '10%'), and maximum amount.
    """
    try:
        # Convert "10%" -> 0.10
        percent_value = float(percentage.strip('%')) / 100
    except ValueError:
        percent_value = 0.10  # Default to 10% if invalid input

    calculated_limit = int(sum_insured * percent_value)
    
    # Return the smaller value between calculated and max cap
    return min(calculated_limit, max_amount)


def get_complete_row_template():
    """Return a complete row template with all required fields as empty strings"""
    return {
        "Max No Of Members Covered": "",
        "Relationship Covered (Member Count)": "",
        "Relationship Covered": "",
        "Min_Age(In Years)": "",
        "Min_Age(In Months)": "",
        "Max_Age(In Years)": "",
        "Max_Age(In Months)": "",
        "Member_Count": "",
        "Member_Type": "",
        "Sublimit_Applicable": "",
        "Sublimit_Type": "",
        "Sub_Limit": "",
        "Family Buffer Applicable": "",
        "Family Buffer Amount": "",
        "Is Network Applicable": "",
        "Black listed hospitals are applicable?": "",
        "Corporate Buffer applicable": "",
        "Buffer Type": "",
        "Applicable for": "",
        "Total Corporate Buffer": "",
        "Corporate Buffer Limit Per Family": "",
        "Corporate Buffer Limit Per Parent": "",
        "Reload of SI": "",
        "Total Corporate Buffer.1": "",
        "Corporate Buffer Limit Per Family.1": "",
        "Corporate Buffer Limit Per Parent.1": "",
        "Reload of SI.1": "",
        "Approving Authority": "",
        "Buffer OPD Limit": "",
        "Whether increase in sum insured permissible at renewal": "",
        "Total Plan Buffer": "",
        "Corporate Bufferr Limit for Employee/Family": "",
        "Corporate Buffer Limit Per Parent.2": "",
        "Reload of SI.2": "",
        "Approving Authority.1": "",
        "Buffer OPD Limit.1": "",
        "Whether increase in sum insured permissible at renewal.1": "",
        "Critical Illness applicable": "",
        "Critical Illness limit per family": "",
        "Critical Illness Approving Authority": "",
        "Critical Illness Whether increase in sum insured permissible at renewal": ""
    }

def extract_Eligibility(text: str) -> List[Dict]:
    # Extract Endorsement No. 1 section
    endorsement_1_match = re.search(r'Endt\.\s*No\.\s*1.*?(?=Endt\.\s*No\.|Endorsement\s*No\.|$)', text, re.IGNORECASE | re.DOTALL)
    
    if not endorsement_1_match:
        # Try alternative patterns for Endt. No. 1
        endorsement_1_match = re.search(r'Endorsement\s*No\.\s*1.*?(?=Endt\.\s*No\.|Endorsement\s*No\.|$)', text, re.IGNORECASE | re.DOTALL)
        
    if not endorsement_1_match:
        # Try to find any section that mentions eligibility or member definitions
        endorsement_1_match = re.search(r'(?:Eligibility|Member|Insured Person).*?(?=Endt\.\s*No\.|Endorsement\s*No\.|$)', text, re.IGNORECASE | re.DOTALL)
    
    if not endorsement_1_match:
        print("[WARNING] Endt. No. 1 section not found, using entire text for extraction")
        endorsement_1_text = text
    else:
        endorsement_1_text = endorsement_1_match.group(0)
    
    # Extract age ranges dynamically from Endorsement No. 1 only
    age_ranges = extract_age_ranges(endorsement_1_text)
    
    # Extract member counts from Endorsement No. 1 only
    # More robust employee detection - look for various patterns that indicate employee coverage
    employee = 0
    if re.search(r'Insured Person covers employees of the Insured', endorsement_1_text, re.IGNORECASE):
        employee = 1
    elif re.search(r'employee', endorsement_1_text, re.IGNORECASE):
        employee = 1
    elif re.search(r'insured person', endorsement_1_text, re.IGNORECASE):
        employee = 1
    # If we have spouse or children but no explicit employee detection, assume employee is covered
    elif re.search(r'spouse', endorsement_1_text, re.IGNORECASE) or re.search(r'children', endorsement_1_text, re.IGNORECASE):
        employee = 1
    
    # More robust spouse detection
    spouse = 0
    spouse_patterns = [
        r'Spouse\s*-\s*Spouse means the employee[\u2019\'s]{0,2}\s+legally married partner',
        r'spouse',
        r'Dependent Spouse',
        r'Dependant Spouse',
        r'Spouse\s*-\s*(\d+)',
        r'Spouse\s*:\s*(\d+)'
    ]
    
    for pattern in spouse_patterns:
        if re.search(pattern, endorsement_1_text, re.IGNORECASE):
            spouse = 1
            break
    
    # Extract children count dynamically from Endorsement No. 1 only
    children = 0
    
    # Multiple patterns to catch different formats of dependent children mentions
    children_patterns = [
        r'Maximum of (?:the )?first\s*(\d+)\s*dependent children',
        r'Maximum of (?:the )?first\s*(\d+)\s*dependant children',
        r'(\d+)\s*dependent children',
        r'(\d+)\s*dependant children',
        r'Maximum of (?:the )?first\s*(\d+)\s*children',
        r'(\d+)\s*children',
        r'Dependent Children\s*-\s*(\d+)',
        r'Dependant Children\s*-\s*(\d+)',
        r'Children\s*-\s*(\d+)',
        r'Maximum of (?:the )?first\s*(\d+)\s*dependent',
        r'(\d+)\s*dependent',
        r'Children\s*:\s*(\d+)',
        r'Dependent Children\s*:\s*(\d+)',
        r'Dependant Children\s*:\s*(\d+)',
        r'Maximum\s*(\d+)\s*children',
        r'Maximum\s*(\d+)\s*dependent',
        r'Maximum\s*(\d+)\s*dependant',
        r'Up to\s*(\d+)\s*children',
        r'Up to\s*(\d+)\s*dependent',
        r'Up to\s*(\d+)\s*dependant'
    ]
    
    for pattern in children_patterns:
        children_match = re.search(pattern, endorsement_1_text, re.IGNORECASE)
        if children_match:
            children = int(children_match.group(1))
            break
    
    # If no specific count found, check if dependent children are mentioned at all
    if children == 0:
        if re.search(r'dependent children', endorsement_1_text, re.IGNORECASE) or re.search(r'dependant children', endorsement_1_text, re.IGNORECASE):
            # Default to 2 children if mentioned but no count specified
            children = 2
    
    # If we have children but no employee detected, assume employee is covered
    if children > 0 and employee == 0:
        employee = 1
        print(f"[INFO] Employee coverage assumed due to {children} dependent children found")
    
    # Extract parents count from Endorsement No. 1 only
    # Check for various parent patterns
    dependent_parents = 0
    parents_in_law = 0
    
    # Patterns for dependent parents
    dependent_parents_patterns = [
        r'Dependent Parents',
        r'Dependant Parents',
        r'Dependent Parent',
        r'Dependant Parent',
        r'Parents\s*-\s*(\d+)',
        r'Parent\s*-\s*(\d+)'
    ]
    
    for pattern in dependent_parents_patterns:
        if re.search(pattern, endorsement_1_text, re.IGNORECASE):
            dependent_parents = 2  # Default to 2 (father + mother)
            break
    
    # Patterns for parents in law
    parents_in_law_patterns = [
        r'Dependant Parents in law',
        r'Dependent Parents in law',
        r'Dependant Parents-in-law',
        r'Dependent Parents-in-law',
        r'Parents in law',
        r'Parents-in-law'
    ]
    
    for pattern in parents_in_law_patterns:
        if re.search(pattern, endorsement_1_text, re.IGNORECASE):
            parents_in_law = 2  # Default to 2 (father-in-law + mother-in-law)
            break
    
    parents = dependent_parents + parents_in_law
    
    # Calculate total members covered
    total_covered = employee + spouse + children + parents
    
    # Debug information
    print(f"[DEBUG] Endt. No. 1 extraction results:")
    print(f"[DEBUG] - Employee: {employee}")
    print(f"[DEBUG] - Spouse: {spouse}")
    print(f"[DEBUG] - Children: {children}")
    print(f"[DEBUG] - Parents: {parents}")
    print(f"[DEBUG] - Total covered: {total_covered}")
    
    # Extract buffer information from Endorsement No. 1 only
    buffer_match = re.search(r'limit of Rs[\.:]?\s?([\d,]+)', endorsement_1_text, re.IGNORECASE)
    buffer_limit = int(buffer_match.group(1).replace(',', '')) if buffer_match else 0

    # Extract Sum Insured from the policy text
    sum_insured = 0
    sum_insured_match = re.search(r'limit of Rs\.([\d,]+)/- as Corporate floater', text, re.IGNORECASE)
    if sum_insured_match:
        sum_insured = int(sum_insured_match.group(1).replace(",", ""))
    else:
        # Try alternative patterns for sum insured
        sum_insured_match = re.search(r'sum insured.*?Rs\.([\d,]+)', text, re.IGNORECASE)
        if sum_insured_match:
            sum_insured = int(sum_insured_match.group(1).replace(",", ""))
        else:
            # Default sum insured if not found
            sum_insured = 500000  # Default value

    # Extract Corporate Buffer values from Endorsement No: 10
    corporate_buffer_applicable = extract_corporate_buffer_applicability(text)
    corporate_buffer_limit_family = 0
    corporate_buffer_limit_parent = 0
    reload_of_si = "No limit for the reload of SI"
    buffer_opd_limit = 0
    
    # Look for Endorsement No: 10
    endorsement_10_match = re.search(r'Endt\.\s*No\.\s*10.*?(?=Endt\.\s*No\.|Endorsement\s*No\.|$)', text, re.IGNORECASE | re.DOTALL)
    if endorsement_10_match:
        endorsement_10_text = endorsement_10_match.group(0)
        
        # Check if Corporate Buffer/Floater is applicable
        if re.search(r'corporate\s+(?:buffer|floater)', endorsement_10_text, re.IGNORECASE):
            corporate_buffer_applicable = "Yes"
            
            # Extract Corporate Buffer/Floater Limit
            floater_limit_match = re.search(r'limit\s+of\s+Rs[\.:]?\s?([\d,]+).*?corporate\s+(?:buffer|floater)', endorsement_10_text, re.IGNORECASE)
            if floater_limit_match:
                corporate_buffer_limit_family = int(floater_limit_match.group(1).replace(',', ''))
            
            # Extract Corporate Buffer Limit Per Family (if different from floater limit)
            family_limit_match = re.search(r'corporate\s+buffer\s+limit\s+per\s+family.*?Rs[\.:]?\s?([\d,]+)', endorsement_10_text, re.IGNORECASE)
            if family_limit_match:
                corporate_buffer_limit_family = int(family_limit_match.group(1).replace(',', ''))
            
            # Extract Corporate Buffer Limit Per Parent
            parent_limit_match = re.search(r'corporate\s+buffer\s+limit\s+per\s+parent.*?Rs[\.:]?\s?([\d,]+)', endorsement_10_text, re.IGNORECASE)
            if parent_limit_match:
                corporate_buffer_limit_parent = int(parent_limit_match.group(1).replace(',', ''))
            
            # Extract Reload of SI options - Per person limit pattern
            if re.search(r'per\s+person\s+limit', endorsement_10_text, re.IGNORECASE):
                equivalent_pattern = r'equivalent\s+to\s+the\s+per\s+person\s+limit'
                double_pattern = r'double\s+to\s+the\s+per\s+person\s+limit'
                thrice_pattern = r'thrice\s+to\s+the\s+per\s+person\s+limit'

                has_equivalent = re.search(equivalent_pattern, endorsement_10_text, re.IGNORECASE)
                has_double = re.search(double_pattern, endorsement_10_text, re.IGNORECASE)
                has_thrice = re.search(thrice_pattern, endorsement_10_text, re.IGNORECASE)

                # Check for specific patterns
                if has_equivalent:
                    reload_of_si = "Reload of SI is up to the Existing SI"
                elif has_double:
                    reload_of_si = "Reload of SI is up to the Double the SI"
                elif has_thrice:
                    reload_of_si = "Reload of SI is up to the Thrice the SI"
                else:
                    reload_of_si = "No limit for the reload of SI"
            
            # Extract Buffer OPD Limit
            opd_limit_match = re.search(r'buffer\s+opd\s+limit.*?Rs[\.:]?\s?([\d,]+)', endorsement_10_text, re.IGNORECASE)
            if opd_limit_match:
                buffer_opd_limit = int(opd_limit_match.group(1).replace(',', ''))

    # Extract Critical Illness values based on Corporate Buffer logic
    critical_illness_limit_family = 0.0
    
    # Determine Critical Illness applicable based on Corporate Buffer status
    if corporate_buffer_applicable == "Yes":
        critical_illness_applicable = ""  # Empty if Corporate Buffer is Yes
    else:
        critical_illness_applicable = "No"  # No if Corporate Buffer is empty/No
    
    # Extract Critical Illness limit per family only if applicable
    if critical_illness_applicable != "No":
        critical_limit_match = re.search(r'critical\s+illness\s+limit\s+per\s+family.*?Rs[\.:]?\s?([\d,]+)', text, re.IGNORECASE)
        if critical_limit_match:
            critical_illness_limit_family = float(critical_limit_match.group(1).replace(',', ''))
        else:
            # Look for general critical illness limit
            general_critical_match = re.search(r'critical\s+illness.*?limit.*?Rs[\.:]?\s?([\d,]+)', text, re.IGNORECASE)
            if general_critical_match:
                critical_illness_limit_family = float(general_critical_match.group(1).replace(',', ''))

    # Extract ALL sublimits from Endt. No. 5(i) and Endt. No. 5(ii)
    all_sublimits_5i = extract_sublimits_from_endorsement_5i(text)
    all_sublimits_5ii = extract_all_sublimits_from_endorsement_5ii(text)
    
    # Combine both sets of sublimits
    all_sublimits = all_sublimits_5i + all_sublimits_5ii
    
    # Debug: Check what sublimits were extracted
    print(f"[DEBUG] Total sublimits extracted from 5(i): {len(all_sublimits_5i)}")
    print(f"[DEBUG] Total sublimits extracted from 5(ii): {len(all_sublimits_5ii)}")
    print(f"[DEBUG] Combined total sublimits: {len(all_sublimits)}")
    for i, sublimit in enumerate(all_sublimits):
        print(f"[DEBUG] Sublimit {i+1}: Type='{sublimit.get('type', 'MISSING')}', Limit='{sublimit.get('limit', 'MISSING')}'")
    
   
    covers = []

    # MAIN EMPLOYEE ROW - Contains all the comprehensive data including first sublimit
    if employee > 0:
        age_data = age_ranges["employee"]
        member_type = determine_member_type(age_data["min_years"], age_data["max_years"])
        
        # Create the main Employee row with ALL comprehensive information INCLUDING first sublimit
        main_employee_row = {
            "Max No Of Members Covered": total_covered,
            "Relationship Covered (Member Count)": total_covered,
            "Relationship Covered": "Employee",
            "Min_Age(In Years)": age_data["min_years"] if age_data["min_years"]  else "0",
            "Min_Age(In Months)": age_data["min_months"] if age_data["min_months"] != 0 else "",
            "Max_Age(In Years)": age_data["max_years"] if age_data["max_years"] != 0 else "",
            "Max_Age(In Months)": age_data["max_months"] if age_data["max_months"] != 0 else "",
            "Member_Count": employee,
            "Member_Type": member_type,
            "Sublimit_Applicable": "Yes" if all_sublimits else "",
            "Sublimit_Type": all_sublimits[0].get("type", "") if all_sublimits and len(all_sublimits) > 0 else "",  # First sublimit type
            "Sub_Limit": str(all_sublimits[0].get("limit", "")) if all_sublimits and len(all_sublimits) > 0 else "",  # First sublimit amount
            "Family Buffer Applicable": corporate_buffer_applicable,
            "Family Buffer Amount": corporate_buffer_limit_family,
            "Is Network Applicable": "No",
            "Black listed hospitals are applicable?": "Yes",

            # Corporate Buffer & Additional Fields
            "Corporate Buffer applicable": "Yes" if corporate_buffer_applicable == "Yes" else "No",
            "Buffer Type": "Both" if corporate_buffer_applicable == "Yes" else "",
            "Applicable for": "",
            "Total Corporate Buffer": "",
            "Corporate Buffer Limit Per Family": "",
            "Corporate Buffer Limit Per Parent": "",
            "Reload of SI": "",
            "Total Corporate Buffer.1": "",
            "Corporate Buffer Limit Per Family.1": "",
            "Corporate Buffer Limit Per Parent.1": "",
            "Reload of SI.1": "",
            "Approving Authority": "" ,
            "Buffer OPD Limit": "",
            "Whether increase in sum insured permissible at renewal": "",
            "Total Plan Buffer": corporate_buffer_limit_family if corporate_buffer_applicable == "Yes" !=0  else "",
            "Corporate Bufferr Limit for Employee/Family": corporate_buffer_limit_family if corporate_buffer_applicable == "Yes" !=0 else "",
            "Corporate Buffer Limit Per Parent.2": corporate_buffer_limit_family if corporate_buffer_applicable == "Yes" !=0 else "",
            "Reload of SI.2": reload_of_si if corporate_buffer_applicable == "Yes" !=0 else "",
            "Approving Authority.1": "Corporate HR" if corporate_buffer_applicable == "Yes" !=0 else "",
            "Buffer OPD Limit.1": buffer_opd_limit if corporate_buffer_applicable == "Yes" and buffer_opd_limit != 0 else "",
            "Whether increase in sum insured permissible at renewal.1": "No" if corporate_buffer_applicable == "Yes" !=0 else "",

            # Critical Illness Fields
            "Critical Illness applicable": critical_illness_applicable,
            "Critical Illness limit per family": "",
            "Critical Illness Approving Authority": "",
            "Critical Illness Whether increase in sum insured permissible at renewal": ""
        }
        covers.append(main_employee_row)

        # SPOUSE ROW - Basic information only
        if spouse > 0:
            age_data = age_ranges["spouse"]
            member_type = determine_member_type(age_data["min_years"], age_data["max_years"])
            
            spouse_row = get_complete_row_template()
            spouse_row.update({
                "Relationship Covered": "Spouse",
                "Min_Age(In Years)": age_data["min_years"] if age_data["min_years"]  else "0",
                "Min_Age(In Months)": age_data["min_months"] if age_data["min_months"] != 0 else "",
                "Max_Age(In Years)": age_data["max_years"] if age_data["max_years"] != 0 else "",
                "Max_Age(In Months)": age_data["max_months"] if age_data["max_months"] != 0 else "",
                "Member_Count": employee,
                "Member_Type": member_type,
                "Sublimit_Applicable": "yes",
                "Sublimit_Type": all_sublimits[1].get("type", "") if all_sublimits and len(all_sublimits) > 1 else "",
                "Sub_Limit": str(all_sublimits[1].get("limit", "")) if all_sublimits and len(all_sublimits) > 1 else ""
            })
            covers.append(spouse_row)

        # SON ROW - Basic information only
        if children > 0:
            age_data = age_ranges["children"]
            member_type = determine_member_type(age_data["min_years"], age_data["max_years"])
            
            son_row = get_complete_row_template()
            son_row.update({
                "Relationship Covered": "Son",
                "Min_Age(In Years)": age_data["min_years"] if age_data["min_years"]  else "0",
                "Min_Age(In Months)": age_data["min_months"] if age_data["min_months"] != 0 else "",
                "Max_Age(In Years)": age_data["max_years"] if age_data["max_years"] != 0 else "",
                "Max_Age(In Months)": age_data["max_months"] if age_data["max_months"] != 0 else "",
                "Member_Count": employee,
                "Member_Type": member_type,
                "Sublimit_Applicable": "yes",
                "Sublimit_Type": all_sublimits[2].get("type", "") if all_sublimits and len(all_sublimits) > 2 else "",
                "Sub_Limit": str(all_sublimits[2].get("limit", "")) if all_sublimits and len(all_sublimits) > 2 else ""
            })
            covers.append(son_row)

        # DAUGHTER ROW - Basic information only
        if children > 0:
            age_data = age_ranges["children"]
            member_type = determine_member_type(age_data["min_years"], age_data["max_years"])
            
            daughter_row = get_complete_row_template()
            daughter_row.update({
                "Relationship Covered": "Daughter",
                "Min_Age(In Years)": age_data["min_years"] if age_data["min_years"]  else "0",
                "Min_Age(In Months)": age_data["min_months"] if age_data["min_months"] != 0 else "",
                "Max_Age(In Years)": age_data["max_years"] if age_data["max_years"] != 0 else "",
                "Max_Age(In Months)": age_data["max_months"] if age_data["max_months"] != 0 else "",
                "Member_Count": employee,
                "Member_Type": member_type,
                "Sublimit_Applicable": "yes",
                "Sublimit_Type": all_sublimits[3].get("type", "") if all_sublimits and len(all_sublimits) > 3 else "",
                "Sub_Limit": str(all_sublimits[3].get("limit", "")) if all_sublimits and len(all_sublimits) > 3 else ""
            })
            covers.append(daughter_row)

        # FATHER ROW - Basic information only (if dependent parents exist)
        if dependent_parents > 0:
            age_data = age_ranges["parents"]
            member_type = determine_member_type(age_data["min_years"], age_data["max_years"])

            father_row = get_complete_row_template()
            father_row.update({
                "Relationship Covered": "Father",
                "Min_Age(In Years)": age_data["min_years"] if age_data["min_years"]  else "0",
                "Min_Age(In Months)": age_data["min_months"] if age_data["min_months"] != 0 else "",
                "Max_Age(In Years)": age_data["max_years"] if age_data["max_years"] != 0 else "",
                "Max_Age(In Months)": age_data["max_months"] if age_data["max_months"] != 0 else "",
                "Member_Count": employee,
                "Member_Type": member_type,
                "Sublimit_Applicable": "yes",
                "Sublimit_Type": all_sublimits[4].get("type", "") if all_sublimits and len(all_sublimits) > 4 else "",
                "Sub_Limit": str(all_sublimits[4].get("limit", "")) if all_sublimits and len(all_sublimits) > 4 else ""
            })
            covers.append(father_row)

            # MOTHER ROW - Basic information only
            mother_row = get_complete_row_template()
            mother_row.update({
                "Relationship Covered": "Mother",
                "Min_Age(In Years)": age_data["min_years"] if age_data["min_years"]  else "0",
                "Min_Age(In Months)": age_data["min_months"] if age_data["min_months"] != 0 else "",
                "Max_Age(In Years)": age_data["max_years"] if age_data["max_years"] != 0 else "",
                "Max_Age(In Months)": age_data["max_months"] if age_data["max_months"] != 0 else "",
                "Member_Count": employee,
                "Member_Type": member_type,
                "Sublimit_Applicable": "yes",
                "Sublimit_Type": all_sublimits[5].get("type", "") if all_sublimits and len(all_sublimits) > 5 else "",
                "Sub_Limit": str(all_sublimits[5].get("limit", "")) if all_sublimits and len(all_sublimits) > 5 else ""
            })
            covers.append(mother_row)

        # FATHER-IN-LAW ROW - Basic information only (if parents in law exist)
        if parents_in_law > 0:
            age_data = age_ranges["parents"]
            member_type = determine_member_type(age_data["min_years"], age_data["max_years"])
            
            father_in_law_row = get_complete_row_template()
            father_in_law_row.update({
                "Relationship Covered": "Father-in-law",
                "Min_Age(In Years)": age_data["min_years"] if age_data["min_years"]  else "0",
                "Min_Age(In Months)": age_data["min_months"] if age_data["min_months"] != 0 else "",
                "Max_Age(In Years)": age_data["max_years"] if age_data["max_years"] != 0 else "",
                "Max_Age(In Months)": age_data["max_months"] if age_data["max_months"] != 0 else "",
                "Member_Count": employee,
                "Member_Type": member_type,
                "Sublimit_Applicable": "yes",
                "Sublimit_Type": all_sublimits[6].get("type", "") if all_sublimits and len(all_sublimits) > 6 else "",
                "Sub_Limit": str(all_sublimits[6].get("limit", "")) if all_sublimits and len(all_sublimits) > 6 else ""
            })
            covers.append(father_in_law_row)

            # MOTHER-IN-LAW ROW - Basic information only
            mother_in_law_row = get_complete_row_template()
            mother_in_law_row.update({
                "Relationship Covered": "Mother-in-law",
                "Min_Age(In Years)": age_data["min_years"] if age_data["min_years"]  else "0",
                "Min_Age(In Months)": age_data["min_months"] if age_data["min_months"] != 0 else "",
                "Max_Age(In Years)": age_data["max_years"] if age_data["max_years"] != 0 else "",
                "Max_Age(In Months)": age_data["max_months"] if age_data["max_months"] != 0 else "",
                "Member_Count": employee,
                "Member_Type": member_type,
                "Sublimit_Applicable": "yes",
                "Sublimit_Type": all_sublimits[7].get("type", "") if all_sublimits and len(all_sublimits) > 7 else "",
                "Sub_Limit": str(all_sublimits[7].get("limit", "")) if all_sublimits and len(all_sublimits) > 7 else ""
            })
            covers.append(mother_in_law_row)

        # Add additional rows for remaining sublimits (if there are more sublimits than relationships)
        # Calculate how many relationship rows we actually created
        relationship_rows_created = 0
        if employee > 0:
            relationship_rows_created += 1  # Employee
        if spouse > 0:
            relationship_rows_created += 1  # Spouse
        if children > 0:
            relationship_rows_created += 2  # Son + Daughter
        if dependent_parents > 0:
            relationship_rows_created += 2  # Father + Mother
        if parents_in_law > 0:
            relationship_rows_created += 2  # Father-in-law + Mother-in-law
        
        # Add rows for remaining sublimits
        if all_sublimits and len(all_sublimits) > relationship_rows_created:
            for i, sublimit in enumerate(all_sublimits[relationship_rows_created:], relationship_rows_created):
                additional_sublimit_row = get_complete_row_template()
                additional_sublimit_row.update({
                    "Sublimit_Type": sublimit.get("type", ""),
                    "Sub_Limit": str(sublimit.get("limit", ""))
                })
                covers.append(additional_sublimit_row)

        # Clean all data to remove None values and replace with empty strings
        cleaned_covers = []
        for cover in covers:
            cleaned_cover = {}
            for key, value in cover.items():
                if value is None:
                    cleaned_cover[key] = ""
                else:
                    cleaned_cover[key] = value
            cleaned_covers.append(cleaned_cover)

        return cleaned_covers



