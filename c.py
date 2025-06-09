import pandas as pd
import fitz  # PyMuPDF
import re
import os

def extract_text_with_pymupdf(pdf_path):
    """
    Extracts all text from a PDF document using PyMuPDF, page by page.
    Returns a list of lists, where each inner list contains dictionaries
    of text elements (text, bbox, approximate font size) for a page,
    along with page dimensions.
    """
    all_page_data = []
    try:
        document = fitz.open(pdf_path)
        for page_num in range(document.page_count):
            page = document.load_page(page_num)
            
            page_height = page.rect.height
            page_y1 = page.rect.y1

            page_dict = page.get_text("dict")
            
            page_elements = []
            for block in page_dict.get("blocks", []):
                for line in block.get("lines", []):
                    line_text = ""
                    min_x, min_y, max_x, max_y = float('inf'), float('inf'), float('-inf'), float('-inf')
                    font_sizes_in_line = []
                    
                    for span in line.get("spans", []):
                        line_text += span.get("text", "")
                        span_bbox = span.get("bbox")
                        if span_bbox:
                            min_x = min(min_x, span_bbox[0])
                            min_y = min(min_y, span_bbox[1])
                            max_x = max(max_x, span_bbox[2])
                            max_y = max(max_y, span_bbox[3])
                        
                        span_font_size = span.get("size")
                        if span_font_size:
                            font_sizes_in_line.append(span_font_size)
                            
                    line_text = line_text.strip()
                    if line_text:
                        font_size_approx = sum(font_sizes_in_line) / len(font_sizes_in_line) if font_sizes_in_line else 0
                        bbox = (min_x, min_y, max_x, max_y)
                        
                        page_elements.append({
                            'page_num': page_num + 1,
                            'text': line_text,
                            'bbox': bbox,
                            'font_size_approx': font_size_approx,
                            'page_height': page_height,
                            'page_y1': page_y1
                        })
            all_page_data.append(page_elements)
    except Exception as e:
        print(f"Error reading PDF with PyMuPDF: {e}")
    return all_page_data


def detect_university_and_extract_data(pdf_path):
    all_extracted_records = []
    all_pages_elements = extract_text_with_pymupdf(pdf_path)

    current_university_name = "نامشخص"
    
    # Pre-compile common regex patterns for efficiency
    # Using specific Unicode characters to avoid syntax errors
    admission_regex = re.compile(r"\u0635\u0631\u0641\u0627\u0020\u0628\u0627\u0020\u0633\u0648\u0627\u0628\u0642\u0020\u062a\u062d\u0635\u06cc\u0644\u06cc", re.IGNORECASE) # صرفا با سوابق تحصیلی
    code_regex = re.compile(r"\b(\d{5})\b") # 5-digit code
    capacity_regex = re.compile(r"\b(\d{1,3}|-)\b") # Capacity
    sex_regex = re.compile(r"(\u0645\u0631\u062f\s*,\s*\u0632\u0646|\u0645\u0631\u062f|\u0632\u0646)") # مرد , زن | مرد | زن
    description_regex = re.compile(r"(\u0641\u0627\u0642\u062f\u0020\u062e\u0648\u0627\u0628\u06af\u0627\u0647|\u0645\u0639\u0631\u0641\u064a\u0020\u0628\u0647\u0020\u062e\u0648\u0627\u0628\u06af\u0627\u0647\u0020\u0647\u0627\u064a\u0020\u062e\u0648\u062f\u06af\u0631\u062f\u0627\u0646|\u062f\u0627\u0631\u0627\u064a\u060C\u062e\u0648\u0627\u0628\u06af\u0627\u0647\u0020\u0645\u0644\u0643\u064a)") # فاقد خوابگاه | معرفی به خوابگاه های خودگردان | دارای خوابگاه ملکی
    
    # University name keywords (more robust)
    uni_keywords = ["\u0645\u0648\u0633\u0633\u0647", "\u062f\u0627\u0646\u0634\u06af\u0627\u0647", "\u062f\u0627\u0646\u0634\u0643\u062f\u0647"] # موسسه, دانشگاه, دانشکده
    
    # Specific Persian string for "ادامه استان" and "استان"
    persian_continue_state = "\u0627\u062f\u0627\u0645\u0647\u0020\u0627\u0633\u062a\u0627\u0646" # ادامه استان
    persian_state = "\u0627\u0633\u062a\u0627\u0646" # استان
    
    # Common header cleanup words
    persian_header_words = [
        "\u0643\u062f\u0631\u0634\u062a\u0647\u0020\u0645\u062d\u0644", # کدرشته محل
        "\u0639\u0646\u0648\u0627\u0646\u0020\u0631\u0634\u062a\u0647", # عنوان رشته
        "\u0646\u062d\u0648\u0647\u0020\u067e\u0630\u06cc\u0631\u0634", # نحوه پذیرش
        "\u062c\u0646\u0633\u0020\u067e\u0630\u06cc\u0631\u0634", # جنس پذیرش
        "\u0638\u0631\u0641\u06cc\u062a", # ظرفیت
        "\u062a\u0648\u0636\u06cc\u062d\u0627\u062a", # توضیحات
        "\u0627\u0648\u0644", # اول
        "\u062f\u0648\u0645", # دوم
        "\u067e\u0630\u06cc\u0631\u0634", # پذیرش
        "\u0645\u062d\u0644", # محل
        "\u062f\u0641\u062a\u0631\u0686\u0647\u0020\u0631\u0627\u0647\u0646\u0645\u0627\u064a\u0020\u0627\u0646\u062a\u06ﺨ\u0627\u0628\u0020\u0631\u0634\u062a\u0647", # دفترچه راهنمای انتخاب رشته
        "\u0622\u0632\u0645\u0648\u0646\u0020\u0633\u0631\u0627\u0633\u0631\u064a\u0020\u0633\u0627\u0644", # آزمون سراسري سال
        "\u0627\u062f\u0627\u0645\u0647" # ادامه (standalone word)
    ]

    for page_elements in all_pages_elements:
        if not page_elements:
            continue

        page_height = page_elements[0].get('page_height', 792)
        page_y1 = page_elements[0].get('page_y1', page_height)

        # 1. Detect University Name for the current page/section
        max_font_size_on_page = 0
        if page_elements:
            valid_font_sizes = [elem['font_size_approx'] for elem in page_elements if isinstance(elem['font_size_approx'], (int, float))]
            if valid_font_sizes:
                max_font_size_on_page = max(valid_font_sizes)
        
        university_candidates = []
        for elem in page_elements:
            text = elem['text']
            font_size = elem['font_size_approx']
            y_pos = elem['bbox'][1]

            is_in_top_section = y_pos > (page_y1 - page_height * 0.3)
            
            if any(keyword in text for keyword in uni_keywords) and \
               (font_size >= max_font_size_on_page * 0.8) and \
               is_in_top_section:
                university_candidates.append(elem)
        
        if university_candidates:
            university_candidates.sort(key=lambda x: x['bbox'][1], reverse=True)
            
            temp_university_name_parts = []
            
            main_candidate_found = False
            for cand in university_candidates:
                # Use the Unicode escaped string here
                if cand['text'].strip().startswith(persian_continue_state):
                    continue 

                if any(keyword in cand['text'] for keyword in uni_keywords):
                    temp_university_name_parts.append(cand['text'].strip())
                    main_candidate_found = True
                    if persian_state in cand['text']:
                        break 
                elif not main_candidate_found and persian_state in cand['text']:
                    temp_university_name_parts.append(cand['text'].strip())
                    break
            
            full_detected_name = " ".join(temp_university_name_parts).replace("  ", " ").strip()
            
            if full_detected_name and \
               any(keyword in full_detected_name for keyword in uni_keywords) and \
               not full_detected_name.startswith(persian_continue_state):
                 current_university_name = full_detected_name

        # 2. Extract Table Data from the rest of the page elements
        for elem in page_elements:
            line_text = elem['text']
            
            # Skip irrelevant lines
            if any(word in line_text for word in persian_header_words) or \
               len(line_text.strip()) < 10 or \
               line_text.strip().isdigit() or \
               line_text.strip() == "\u0627\u062f\u0627\u0645\u0647": # "ادامه" (standalone)
                continue

            # Check if this line is likely a table data row by finding key identifiers
            # We must find admission method AND a 5-digit code in the line
            if admission_regex.search(line_text) and code_regex.search(line_text):
                
                record = {
                    "نام دانشگاه": current_university_name,
                    "نحوه پذیرش": "\u0635\u0631\u0641\u0627\u0020\u0628\u0627\u0020\u0633\u0648\u0627\u0628\u0642\u0020\u062a\u062d\u0635\u06cc\u0644\u06cc", # "صرفا با سوابق تحصیلی"
                    "کدرشته محل": "",
                    "عنوان رشته": "",
                    "ظرفیت": "",
                    "جنس": "",
                    "توضیحات": ""
                }
                
                # Extract Code ID
                code_match = code_regex.search(line_text)
                if code_match:
                    record["کدرشته محل"] = code_match.group(1).strip()
                    line_text = line_text.replace(code_match.group(0), '', 1).strip()
                
                # Extract Capacity
                capacity_match = capacity_regex.search(line_text)
                if capacity_match:
                    record["ظرفیت"] = capacity_match.group(1).strip()
                    line_text = line_text.replace(capacity_match.group(0), '', 1).strip()

                # Extract Sex
                sex_match = sex_regex.search(line_text)
                if sex_match:
                    record["جنس"] = sex_match.group(0).strip()
                    line_text = line_text.replace(sex_match.group(0), '', 1).strip()

                # Extract Description
                description_match = description_regex.search(line_text)
                if description_match:
                    record["توضیحات"] = description_match.group(0).strip()
                    line_text = line_text.replace(description_match.group(0), '', 1).strip()

                # The remaining text should largely be the Course Title.
                cleaned_course_title = re.sub(r'\s*\d+\s*', ' ', line_text).strip()
                
                # Remove common table headers that might be leftover in the course title
                # Using the list of Persian header words for clean up
                for word_to_remove in persian_header_words:
                    cleaned_course_title = cleaned_course_title.replace(word_to_remove, '').strip()

                cleaned_course_title = re.sub(r'\s+', ' ', cleaned_course_title).strip()
                
                record["عنوان رشته"] = cleaned_course_title
                
                if record["کدرشته محل"]:
                    all_extracted_records.append(record)

    return all_extracted_records

# --- Main execution ---
if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    pdf_file_name = "a.pdf"
    pdf_input_path = os.path.join(script_dir, pdf_file_name)
    
    output_excel_path = os.path.join(script_dir, 'extracted_university_courses.xlsx')

    print(f"Starting extraction from {pdf_input_path}...")
    extracted_data = detect_university_and_extract_data(pdf_input_path)

    if extracted_data:
        df = pd.DataFrame(extracted_data)
        
        final_columns = [
            "نام دانشگاه",
            "نحوه پذیرش",
            "کدرشته محل",
            "عنوان رشته",
            "ظرفیت",
            "جنس",
            "توضیحات"
        ]
        
        for col in final_columns:
            if col not in df.columns:
                df[col] = ''
        
        df = df[final_columns]

        df.to_excel(output_excel_path, index=False, engine='openpyxl')
        print(f"Data extracted and saved to {output_excel_path}")
    else:
        print("No data extracted. Please check the PDF structure and regex patterns.")