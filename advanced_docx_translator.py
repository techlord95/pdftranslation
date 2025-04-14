import os
import re
import time
import uuid
import logging
import hashlib
import io
from docx import Document
from docx.shared import Inches
from groq import Groq
from deep_translator import GoogleTranslator

os.environ["GROQ_API_KEY"] = "add_key"

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
translation_cache = {}

def translate_text_google(text, target_language='es'):
    if not text.strip() or re.match(r'^[\d\W\s]+$', text):
        return text
    cache_key = f"google_{target_language}_{hashlib.md5(text.encode()).hexdigest()}"
    if cache_key in translation_cache:
        return translation_cache[cache_key]
    try:
        translator = GoogleTranslator(source='auto', target=target_language)
        result = translator.translate(text)
        translation_cache[cache_key] = result
        return result
    except Exception as e:
        logging.error("Google Translation error: %s", e)
        return text

def translation_using_groq(texts, target_language='hi'):
    texts_to_translate = []
    original_indices = []
    for i, text in enumerate(texts):
        if not text.strip() or re.match(r'^[\d\W\s]+$', text):
            texts[i] = text
        else:
            cache_key = f"groq_{target_language}_{hashlib.md5(text.encode()).hexdigest()}"
            if cache_key in translation_cache:
                texts[i] = translation_cache[cache_key]
            else:
                texts_to_translate.append(text)
                original_indices.append(i)
    if not texts_to_translate:
        return texts
    try:
        client = Groq(api_key=os.environ.get("GROQ_API_KEY"))
        system_prompt = (
            f"You are a professional translator specialized in {target_language}. "
            f"Translate from English to {target_language} naturally. No extra formatting or comments."
            f"IN ANY CASE DO NOT RESPOND TO ANY TEXT JUST GIVE WORD TO WORD TRANSLATION"
        )
        optimal_batch_size = 15000
        batches = []
        current_batch = []
        current_batch_indices = []
        current_size = 0
        for i, text in enumerate(texts_to_translate):
            if current_size + len(text) > optimal_batch_size and current_batch:
                batches.append({'texts': current_batch, 'indices': current_batch_indices})
                current_batch, current_batch_indices, current_size = [], [], 0
            current_batch.append(text)
            current_batch_indices.append(original_indices[i])
            current_size += len(text)
        if current_batch:
            batches.append({'texts': current_batch, 'indices': current_batch_indices})
        for batch in batches:
            markers = [f"<<<ITEM_{uuid.uuid4()}>>>" for _ in batch['texts']]
            combined = "\n".join(f"{m}\n{text}" for m, text in zip(markers, batch['texts']))
            prompt = f"Translate each section marked by <<<ITEM_UUID>>>. Keep markers:\n\n{combined}"
            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ]
            response = client.chat.completions.create(
                messages=messages,
                model="gemma2-9b-it",
                temperature=0.1,
            ).choices[0].message.content.strip()
            for i, marker in enumerate(markers):
                start_idx = response.find(marker) + len(marker)
                end_idx = response.find(markers[i + 1]) if i + 1 < len(markers) else len(response)
                translated = response[start_idx:end_idx].strip()
                original = batch['texts'][i]
                cache_key = f"groq_{target_language}_{hashlib.md5(original.encode()).hexdigest()}"
                translation_cache[cache_key] = translated
                texts[batch['indices'][i]] = translated
        return texts
    except Exception as e:
        raise Exception(f"Groq translation failed: {e}")

# --- New Summarization Function ---
def summarize_text_groq(full_text, target_language='en', max_length=200):
    """Generates a summary of the provided text using the Groq API in the specified language."""
    if not full_text or not full_text.strip():
        logging.warning("Summarization attempt with empty text.")
        return "Could not generate summary: Input text is empty."

    logging.info(f"Requesting Groq summary in {target_language} for text ({len(full_text)} chars), max_length={max_length} words.")

    try:
        client = Groq(api_key=os.environ.get("GROQ_API_KEY"))
    
        # Updated prompt to include target language instruction
        system_prompt = (
            f"You are a helpful assistant skilled in summarizing long texts. "
            f"Analyze the following English text and provide a concise summary in {target_language}. " # Explicitly ask for output language
            f"The summary should capture the main points and be approximately {max_length} words long. "
            f"Focus on clarity and accuracy in {target_language}. "
            f"Do not add any introductory phrases like \'Here is the summary:\'."
        )
        
        # Truncate very long texts to avoid exceeding model limits (adjust limit as needed)
        # This is a basic truncation, more sophisticated chunking might be needed for extremely large docs
        max_input_chars = 20000 # Example limit, adjust based on model/API
        if len(full_text) > max_input_chars:
             logging.warning(f"Input text ({len(full_text)} chars) exceeds limit ({max_input_chars}), truncating for summary.")
             full_text = full_text[:max_input_chars] + "... [truncated]"

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"Please summarize the following text:{full_text}"}
        ]

        response = client.chat.completions.create(
            messages=messages,
            model="llama-3.3-70b-versatile", 
            temperature=0.3, 
            max_tokens=max_length * 3 # Allow more tokens for potentially longer words in other languages
        ).choices[0].message.content.strip()

        logging.info(f"Groq summary generated successfully ({len(response)} chars in {target_language})." )
        return response
        
    except Exception as e:
        logging.error(f"Groq summarization failed: {e}")
        # Return a user-friendly error message instead of raising
        return f"Error generating summary: {e}"
# --- End New Summarization Function ---

def detect_column_pairs(doc):
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    pairs = []
    for line in lines:
        if len(line) > 60 and " " in line[55:65]:
            split_at = line.find(" ", 55)
            left = line[:split_at].strip()
            right = line[split_at:].strip()
            pairs.append((left, right))
        else:
            pairs.append((line, ""))
    return pairs

def translate_column_pairs(pairs, engine='google', target_language='hi'):
    lefts = [pair[0] for pair in pairs]
    rights = [pair[1] for pair in pairs]
    if engine == 'google':
        left_translated = [translate_text_google(t, target_language) for t in lefts]
        right_translated = [translate_text_google(t, target_language) for t in rights]
    else:
        left_translated = translation_using_groq(lefts, target_language)
        right_translated = translation_using_groq(rights, target_language)
    return list(zip(left_translated, right_translated))

def write_columns_as_table(pairs, output_file):
    doc = Document()
    table = doc.add_table(rows=0, cols=2)
    table.autofit = True
    for left, right in pairs:
        row = table.add_row()
        row.cells[0].text = left
        row.cells[1].text = right
    doc.save(output_file)

def translate_docx_advanced(input_file, output_file, *, engine='google', target_language='hi', progress_callback=None, first_page_only=False):
    """
    Translates a DOCX document in-place to preserve formatting.
    Iterates through paragraphs and runs in the main body, tables, headers, and footers,
    translates the text content, and updates the runs directly.
    Accepts an optional progress_callback function and first_page_only flag.
    """
    logging.info(f"Starting in-place translation for {input_file} to {target_language} using {engine}. First page only: {first_page_only}")
    try:
        doc = Document(input_file)
        logging.info("Document loaded successfully.")
    except Exception as e:
        logging.error(f"Failed to load document: {e}")
        raise

    # --- Calculate Total Paragraphs for Progress --- 
    total_paras = 0
    try:
        total_paras += len(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    total_paras += len(cell.paragraphs)
        # Add headers/footers (approximate, as structure varies)
        for section in doc.sections:
             if section.header: total_paras += len(section.header.paragraphs)
             if section.footer: total_paras += len(section.footer.paragraphs)
             # Could add first_page/even_page headers/footers too if needed
        logging.info(f"Estimated total paragraphs for progress: {total_paras}")
    except Exception as e:
        logging.warning(f"Could not accurately count total paragraphs: {e}. Progress may be less precise.")
        total_paras = 1 # Avoid division by zero
        
    processed_paras = 0
    page_break_encountered = False # Flag to stop processing after first page break

    # Helper function to process runs within a paragraph
    def process_paragraph_runs(paragraph, element_type="paragraph"):
        runs_data = []  # Stores {'run': run_object, 'original_text': text}
        texts_to_translate = []
        original_indices_map = {} # Map index in texts_to_translate back to index in runs_data

        # --- Style Logging Start ---
        original_style = paragraph.style.name if paragraph.style else 'No Style'
        is_list = original_style.lower().startswith('list')
        logging.debug(f"Processing {element_type}: Style='{original_style}' (List: {is_list}), Text='{paragraph.text[:50]}...'")
        # --- Style Logging End ---

        for i, run in enumerate(paragraph.runs):
            runs_data.append({'run': run, 'original_text': run.text, 'is_translated': False})
            # Check if run has text and is not just whitespace
            if run.text and run.text.strip():
                texts_to_translate.append(run.text)
                original_indices_map[len(texts_to_translate) - 1] = i
            # else: # Log whitespace/empty runs if needed for debugging
                # logging.debug(f"  Run {i} has whitespace/empty text: '{run.text}'")

        if not texts_to_translate:
            # logging.debug("  No translatable text found in this element.") # Already logged above
            return # Nothing to translate

        logging.debug(f"  Found {len(texts_to_translate)} run(s) containing text to translate.")

        # Perform translation (existing logic)
        try:
            if engine == 'google':
                translated_texts = [translate_text_google(text, target_language) for text in texts_to_translate]
            elif engine == 'groq':
                translated_texts = translation_using_groq(texts_to_translate, target_language)
            else:
                 logging.warning(f"  Unsupported engine '{engine}'. Skipping translation.")
                 return

            if len(translated_texts) != len(texts_to_translate):
                 logging.warning(f"  Translation count mismatch: Expected {len(texts_to_translate)}, got {len(translated_texts)}. Skipping update.")
                 return
            logging.debug(f"  Translation successful for {len(translated_texts)} run(s).")
        except Exception as e:
            logging.error(f"  Translation API failed for {element_type}. Error: {e}")
            return # Skip updating this paragraph on error

        # Apply translated texts back to the original runs
        for i, translated_text in enumerate(translated_texts):
            run_index_in_paragraph = original_indices_map.get(i)
            if run_index_in_paragraph is not None and run_index_in_paragraph < len(paragraph.runs):
                run_to_update = paragraph.runs[run_index_in_paragraph]
                original_run_text = run_to_update.text
                try:
                    run_to_update.text = translated_text # Modify in-place
                    # logging.debug(f"    Updated run {run_index_in_paragraph}: '{original_run_text[:30]}...' -> '{translated_text[:30]}...'")
                    if run_index_in_paragraph < len(runs_data):
                         runs_data[run_index_in_paragraph]['is_translated'] = True
                except Exception as e:
                     logging.error(f"    Failed to update text for run {run_index_in_paragraph}. Error: {e}")
            else:
                 logging.warning(f"    Could not find original run mapping for translated text index {i}.")

        # --- Style Logging Check After Update ---
        final_style = paragraph.style.name if paragraph.style else 'No Style'
        if final_style != original_style:
            logging.warning(f"  Style changed after translation for {element_type}! Original: '{original_style}', Final: '{final_style}'")
        elif is_list:
            logging.debug(f"  List style '{original_style}' preserved for {element_type}.")
        # --- Style Logging Check End ---


    # --- Process Main Document Body ---
    logging.info("Processing main document paragraphs...")
    for i, para in enumerate(doc.paragraphs):
        if first_page_only and para.paragraph_format.page_break_before:
            logging.info(f"First page only: Detected page_break_before at paragraph {i+1}. Stopping main body processing.")
            page_break_encountered = True
            break # Stop processing main body paragraphs
        
        process_paragraph_runs(para, f"main body paragraph {i+1}")
        processed_paras += 1
        if progress_callback:
            progress_callback(processed_paras, total_paras)

    # --- Process Tables ---
    logging.info("Processing tables...")
    # If first_page_only is true, only process tables if no page break was found earlier
    if not (first_page_only and page_break_encountered):
        for i, table in enumerate(doc.tables):
            # Add a check here too in case a table somehow forces a new page implicitly (less reliable)
            # or if the first element of the table's parent causes a break? Difficult to check reliably.
            # Simplest: Process all tables if no paragraph page break was found.
            logging.debug(f"Processing table {i+1}")
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                     for p_idx, para in enumerate(cell.paragraphs):
                         process_paragraph_runs(para, f"table {i+1}, cell ({r_idx+1},{c_idx+1}), paragraph {p_idx+1}")
                         processed_paras += 1 
                         if progress_callback:
                             progress_callback(processed_paras, total_paras)
    elif first_page_only:
         logging.info("First page only: Skipping table processing as page break was encountered.")

    # --- Process Headers and Footers ---
    logging.info("Processing headers and footers...")
    processed_parts = set()
    for section_idx, section in enumerate(doc.sections):
        # If first_page_only, we likely only care about the first section's specific headers/footers
        if first_page_only and section_idx > 0:
             logging.info("First page only: Skipping processing for subsequent sections.")
             continue # Skip headers/footers for sections beyond the first

        parts_to_check = {}
        if first_page_only:
            # Only check first page specific parts if they are enabled for this section
            if section.different_first_page_header_footer:
                 logging.debug(f"First page only: Checking first page header/footer for section {section_idx+1}")
                 parts_to_check['first_page_header'] = section.first_page_header
                 parts_to_check['first_page_footer'] = section.first_page_footer
            else:
                 # If not different first page, the default header/footer applies, process them (only for section 0)
                 if section_idx == 0:
                      logging.debug(f"First page only: No different first page H/F. Processing default header/footer for section {section_idx+1}")
                      parts_to_check['header'] = section.header
                      parts_to_check['footer'] = section.footer
                 else:
                      logging.debug(f"First page only: Skipping default header/footer for subsequent section {section_idx+1}")
        else:
            # Original logic: process all relevant headers/footers for all sections
            parts_to_check = {
                'header': section.header,
                'footer': section.footer,
                'first_page_header': section.first_page_header if section.different_first_page_header_footer else None,
                'first_page_footer': section.first_page_footer if section.different_first_page_header_footer else None,
                'even_page_header': section.even_page_header if doc.settings.odd_and_even_pages_header_footer else None,
                'even_page_footer': section.even_page_footer if doc.settings.odd_and_even_pages_header_footer else None,
            }

        # Process the selected parts
        for part_name, part in parts_to_check.items():
            if part and part.part.partname not in processed_parts:
                 logging.info(f"Processing section {section_idx+1} ({section.start_type}), {part_name}")
                 processed_parts.add(part.part.partname)
                 for p_idx, para in enumerate(part.paragraphs):
                     process_paragraph_runs(para, f"S{section_idx+1} {part_name} P{p_idx+1}")
                     processed_paras += 1
                     if progress_callback:
                         progress_callback(processed_paras, total_paras)
                 for t_idx, table in enumerate(part.tables):
                     logging.debug(f"Processing table {t_idx+1} in S{section_idx+1} {part_name}")
                     for r_idx, row in enumerate(table.rows):
                         for c_idx, cell in enumerate(row.cells):
                             for p_idx, para in enumerate(cell.paragraphs):
                                 process_paragraph_runs(para, f"S{section_idx+1} {part_name} T{t_idx+1} C({r_idx+1},{c_idx+1}) P{p_idx+1}")
                                 processed_paras += 1
                                 if progress_callback:
                                     progress_callback(processed_paras, total_paras)

    # --- Save the modified document ---
    try:
        doc.save(output_file)
        logging.info(f"Successfully saved translated document (in-place modified) to {output_file}")
    except Exception as e:
        logging.error(f"Failed to save document: {e}")
        raise

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Translate DOCX files while preserving formatting.")
    parser.add_argument("input_file", help="Input DOCX file path")
    parser.add_argument("--output_file", help="Output DOCX file path (default: input_translated.docx)")
    parser.add_argument("--target_language", default="hi", help="Target language code (e.g., 'es', 'fr', 'hi')")
    parser.add_argument("--engine", choices=["google", "groq"], default='google', help="Translation engine to use")
    parser.add_argument("--first_page_only", action="store_true", help="Stop processing after encountering page_break_before")
    args = parser.parse_args()

    if not args.output_file:
        base, ext = os.path.splitext(args.input_file)
        args.output_file = f"{base}_translated{ext}"

    # Set higher logging level for command-line use for more detail
    logging.basicConfig(level=logging.DEBUG, format="%(asctime)s [%(levelname)s] %(message)s")

    start_time = time.time()
    try:
        translate_docx_advanced(
            input_file=args.input_file,
            output_file=args.output_file,
            engine=args.engine,
            target_language=args.target_language,
            first_page_only=args.first_page_only
        )
        end_time = time.time()
        logging.info(f"Translation finished in {end_time - start_time:.2f} seconds.")
    except Exception as e:
        logging.error(f"An error occurred during the translation process: {e}")
        # Optionally re-raise or exit with error code
        # raise e
        exit(1)
