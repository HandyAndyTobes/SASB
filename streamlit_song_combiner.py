
import streamlit as st
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO

st.set_page_config(page_title="SASB Song Combiner", layout="wide")
st.title("🎶 SASB Song Combiner Tool")

# Form Inputs
with st.form("song_form"):
    song_nums = st.text_input("Enter Song Numbers (comma-separated)", "2, 3, 5")
    font_color = st.color_picker("Font Color", "#FFFFFF")
    bg_color = st.color_picker("Background Color", "#000000")
    bg_image = st.file_uploader("Upload Background Image (optional)", type=["png", "jpg", "jpeg"])
    logo_image = st.file_uploader("Upload Logo (optional)", type=["png", "jpg", "jpeg"])
    submit = st.form_submit_button("Generate PowerPoint")

def split_text_into_chunks(lines, chunk_size):
    return [lines[i:i + chunk_size] for i in range(0, len(lines), chunk_size)]

def create_combined_pptx(song_numbers, font_color_hex, bg_color_hex, bg_img_bytes, logo_bytes):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    for num in song_numbers:
        match = next((f for f in os.listdir("songs") if f.startswith(f"{num} ")), None)
        if not match:
            continue

        src = Presentation(os.path.join("songs", match))
        for slide in src.slides:
            lines = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        lines.append(para.text.strip())
            if not lines:
                continue

            footer = lines[-1] if len(lines) > 1 else ""
            chunks = split_text_into_chunks(lines[:-1], 8)

            for chunk in chunks:
                s = prs.slides.add_slide(prs.slide_layouts[6])

                # Background color
                s.background.fill.solid()
                s.background.fill.fore_color.rgb = int(bg_color_hex[1:], 16).to_bytes(3, "big")

                if bg_img_bytes:
                    s.shapes.add_picture(bg_img_bytes, 0, 0, width=prs.slide_width, height=prs.slide_height)

                textbox = s.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(12.33), Inches(5.5))
                tf = textbox.text_frame
                tf.clear()
                tf.vertical_anchor = 1  # Middle

                for line in chunk:
                    p = tf.add_paragraph()
                    run = p.add_run()
                    run.text = line
                    run.font.size = Pt(44)
                    run.font.name = 'Calibri'
                    run.font.bold = True
                    run.font.color.rgb = int(font_color_hex[1:], 16).to_bytes(3, "big")
                    p.alignment = 1  # Center

                footer_box = s.shapes.add_textbox(Inches(1), Inches(6.9), Inches(10), Inches(0.5))
                footer_tf = footer_box.text_frame
                footer_tf.text = footer
                footer_tf.paragraphs[0].runs[0].font.size = Pt(20)
                footer_tf.paragraphs[0].runs[0].font.name = "Calibri"
                footer_tf.paragraphs[0].runs[0].font.bold = True
                footer_tf.paragraphs[0].runs[0].font.color.rgb = int(font_color_hex[1:], 16).to_bytes(3, "big")
                footer_tf.paragraphs[0].alignment = 1

                if logo_bytes:
                    s.shapes.add_picture(logo_bytes, Inches(0.2), Inches(6.7), width=Inches(1.0))

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

if submit:
    song_list = [s.strip() for s in song_nums.split(",") if s.strip().isdigit()]
    bg_img_bytes = bg_image if bg_image is not None else None
    logo_bytes = logo_image if logo_image is not None else None
    pptx_file = create_combined_pptx(song_list, font_color, bg_color, bg_img_bytes, logo_bytes)

    st.success("✅ PowerPoint created!")
    st.download_button("Download Presentation", pptx_file, file_name="combined_songs.pptx")
