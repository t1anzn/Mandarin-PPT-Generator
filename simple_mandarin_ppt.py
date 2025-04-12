# -*- coding: utf-8 -*-
import csv
import os
from pptx import Presentation
from pypinyin import pinyin, Style
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from gtts import gTTS
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Function to create a PowerPoint presentation from a template
def create_ppt_from_template(vocab_list, template_path, output_path):
    """
    Create a PowerPoint presentation using a pre-configured template.

    Args:
        vocab_list (list of tuples): List of (Chinese, English) pairs.
        template_path (str): Path to the PowerPoint template file.
        output_path (str): Path to save the generated PowerPoint file.
    """
    # Debugging: Log the path of the template file
    print(f"Loading template from: {template_path}")

    # Debugging: Log the modification timestamp of the template file
    if os.path.exists(template_path):
        modification_time = os.path.getmtime(template_path)
        print(f"Template last modified: {modification_time} (Unix timestamp)")

    # Load the template
    prs = Presentation(template_path)

    # Get the first slide in the template
    template_slide = prs.slides[0]

    # Debugging: Log all placeholders in the template slide
    print("Template slide placeholders:")
    for placeholder in template_slide.placeholders:
        print(f"  idx: {placeholder.placeholder_format.idx}, name: {placeholder.name}, type: {placeholder.placeholder_format.type}")

    # Debugging: Log all shapes in the template slide to identify image placeholders
    print("Template slide shapes:")
    for shape in template_slide.shapes:
        if shape.is_placeholder:
            print(f"  Placeholder idx: {shape.placeholder_format.idx}, name: {shape.name}, type: {shape.placeholder_format.type}")
        else:
            print(f"  Shape: {shape.name}, not a placeholder")

    for chinese, english in vocab_list:
        # Duplicate the template slide
        slide = prs.slides.add_slide(template_slide.slide_layout)

        # Populate placeholders using their idx values
        for placeholder in slide.placeholders:
            if placeholder.placeholder_format.idx == 0:  # Title 1 for CHINESE_PLACEHOLDER
                placeholder.text = chinese
            elif placeholder.placeholder_format.idx == 1:  # Subtitle 2 for ENGLISH_PLACEHOLDER
                placeholder.text = english
            elif placeholder.placeholder_format.idx == 14:  # Updated idx for PINYIN_PLACEHOLDER
                placeholder.text = ' '.join([p[0] for p in pinyin(chinese, style=Style.TONE)])
            elif placeholder.placeholder_format.idx == 15:  # Media Placeholder for SOUND_ICON_PLACEHOLDER
                try:
                    # Generate TTS audio for the Chinese word
                    tts = gTTS(chinese, lang='zh')
                    audio_path = f"media/{chinese}.mp3"
                    tts.save(audio_path)

                    # Debugging: Check if the audio file was created
                    if os.path.exists(audio_path):
                        print(f"Audio file '{audio_path}' successfully created.")
                    else:
                        print(f"Failed to create audio file '{audio_path}'.")

                    # Use the position of the media placeholder
                    left = placeholder.left
                    top = placeholder.top

                    # Add the audio as a media shape to the slide
                    slide.shapes.add_movie(audio_path, left, top, width=placeholder.width, height=placeholder.height)
                    print(f"Added audio '{audio_path}' to the slide at the position of the media placeholder for '{chinese}'.")
                except Exception as e:
                    print(f"Error adding audio for '{chinese}': {e}")

    # Remove the original template slide safely
    xml_slides = prs.slides._sldIdLst  # Access the slide ID list
    slides = list(xml_slides)  # Convert to a list for iteration
    xml_slides.remove(slides[0])  # Remove the first slide (template slide)

    # Save the presentation
    prs.save(output_path)
    print(f"Presentation saved to {output_path}")

if __name__ == "__main__":
    vocab = []
    input_method = input("Enter '1' to load vocabulary from a CSV file or '2' to input manually: ").strip()

    if input_method == '1':
        csv_file = input("Enter the path to the CSV file (default: example_vocab.csv): ").strip()
        if not csv_file:
            csv_file = "example_vocab.csv"

        try:
            with open(csv_file, mode="r", encoding="utf-8") as file:
                reader = csv.reader(file)
                for row in reader:
                    if len(row) >= 2:
                        chinese, english = row[0].strip(), row[1].strip()
                        vocab.append((chinese, english))
        except FileNotFoundError:
            print(f"Error: File '{csv_file}' not found. Exiting.")
            exit(1)
    elif input_method == '2':
        print("Paste vocabulary pairs in the format 'Chinese, English', one pair per line.")
        print("Type 'done' on a new line when you are finished.")
        bulk_input = []
        while True:
            line = input().strip()
            if line.lower() == 'done':
                break
            bulk_input.append(line)

        for line in bulk_input:
            if ',' in line:
                chinese, english = map(str.strip, line.split(',', 1))
                vocab.append((chinese, english))
            else:
                print(f"Skipping invalid line: {line}. Please use 'Chinese, English' format.")
    else:
        print("Invalid option. Exiting.")
        exit(1)

    if not vocab:
        print("No vocabulary provided. Exiting.")
    else:
        template_path = input("Enter the path to the PowerPoint template file (default: template.pptx): ").strip()
        if not template_path:
            template_path = "template.pptx"

        output_path = input("Enter output file path (default: Mandarin_Vocabulary_PPT.pptx): ").strip()
        if not output_path:
            output_path = "Mandarin_Vocabulary_PPT.pptx"
        elif not output_path.lower().endswith(".pptx"):
            output_path += ".pptx"

        create_ppt_from_template(vocab, template_path, output_path)