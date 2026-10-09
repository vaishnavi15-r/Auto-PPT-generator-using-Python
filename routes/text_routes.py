from flask import Blueprint, render_template, request, send_file, session
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer
from collections import Counter
import io
import re
import base64


used_titles = set()

text_bp = Blueprint("text_bp", __name__)

# ----------------------------
# THEME COLORS
# ----------------------------
def get_theme_colors(theme):
    themes = {
        "light": (255, 255, 255),
        "pastel": (223, 239, 255),
        "mint": (230, 249, 230),
        "peach": (255, 229, 180),
        "lavender": (243, 229, 245),
        "sky": (225, 245, 254),
        "gray": (245, 245, 245),
    }
    return themes.get(theme, (255, 255, 255))


# ----------------------------
# SMART SUMMARIZATION
# ----------------------------
def summarize_text(text, sentence_count=8):
    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary = summarizer(parser.document, sentence_count)

    bullets = []
    for sentence in summary:
        bullets.append(str(sentence).strip())

    return bullets


# ----------------------------
# SMART TITLE GENERATOR
# ----------------------------
def generate_slide_title(chunk, full_content):

    global used_titles

    chunk_text = " ".join(chunk)
    chunk_text = re.sub(r'[^a-zA-Z\s]', '', chunk_text).lower()
    full_text = re.sub(r'[^a-zA-Z\s]', '', full_content).lower()

    stopwords = {
        "the","is","are","was","were","and","or","of","to","in","on","for",
        "with","this","that","these","those","it","as","by","an","a",
        "we","you","they","be","from","at","which","into","can","will",
        "has","have","had","their","there","such"
    }

    chunk_words = [
        word for word in chunk_text.split()
        if word not in stopwords and len(word) > 3
    ]

    if not chunk_words:
        return "Overview"

    chunk_counts = Counter(chunk_words)
    full_counts = Counter(full_text.split())

    scored_words = []

    for word, count in chunk_counts.items():
        score = count + full_counts[word] * 0.3
        scored_words.append((word, score))

    scored_words.sort(key=lambda x: x[1], reverse=True)

    title_words = []
    for word, _ in scored_words:
        clean_word = word.capitalize()
        if clean_word not in title_words:
            title_words.append(clean_word)
        if len(title_words) == 2:
            break

    if not title_words:
        return "Overview"

    title = " ".join(title_words)

    if title in used_titles:
        title = title + " Concepts"

    used_titles.add(title)

    return title


# ----------------------------
# PAGE NUMBER
# ----------------------------
def add_page_number(slide, number):
    box = slide.shapes.add_textbox(Inches(9), Inches(6.8), Inches(1), Inches(0.3))
    tf = box.text_frame
    tf.clear()

    p = tf.paragraphs[0]
    p.text = str(number)
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = RGBColor(120, 120, 120)


# ----------------------------
# GENERATE PREVIEW
# ----------------------------
@text_bp.route("/generate/text", methods=["POST"])
def generate_text():

    global used_titles
    used_titles = set()
    
     # --- 1️⃣ Read form inputs ---
    title = request.form.get("title")
    owner = request.form.get("owner")
    content = request.form.get("content")
    font_size = int(request.form.get("font"))
    points_per_slide = int(request.form.get("points"))
    theme = request.form.get("theme")

    # --- ADD THIS: handle user-uploaded image ---
    image_file = request.files.get("slide_image")

    image_bytes = None
    preview_image = None

    if image_file and image_file.filename != "":
        image_bytes = image_file.read()
        preview_image = base64.b64encode(image_bytes).decode("utf-8")


     # --- 3️⃣ Summarize text ---
    lines = summarize_text(content)


    # --- 4️⃣ Store everything in session including image ---
    session["ppt_data"] = {
        "title": title,
        "owner": owner,
        "content": content,
        "font_size": font_size,
        "points_per_slide": points_per_slide,
        "theme": theme,
        "summary_lines": lines
    }

    # --- 5️⃣ Prepare preview slides ---
    bg_r, bg_g, bg_b = get_theme_colors(theme)
    bg_color = f"rgb({bg_r},{bg_g},{bg_b})"


    slides_content = [
        lines[i:i + points_per_slide]
        for i in range(0, len(lines), points_per_slide)
    ]

    preview_slides = []
    preview_titles = []

    for chunk in slides_content:

        slide_title = generate_slide_title(chunk, content)
        preview_titles.append(slide_title)
        preview_slides.append(chunk)

    return render_template(
        "text.html",
        preview=True,
        slides=preview_slides,
        slide_titles=preview_titles,
        bg_color=bg_color,
         preview_image=preview_image
    )


# ----------------------------
# DOWNLOAD PPT
# ----------------------------
@text_bp.route("/download", methods=["GET"])
def download_file():

    global used_titles
    used_titles = set()

    ppt_data = session.get("ppt_data")
    if not ppt_data:
        return "Generate PPT first."

    title = ppt_data["title"]
    owner = ppt_data["owner"]
    content = ppt_data["content"]
    font_size = ppt_data["font_size"]
    points_per_slide = ppt_data["points_per_slide"]
    theme = ppt_data["theme"]
    lines = ppt_data["summary_lines"]

    bg_r, bg_g, bg_b = get_theme_colors(theme)

    prs = Presentation()

    # TITLE SLIDE
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = f"By {owner}"

    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(bg_r, bg_g, bg_b)

    slides_content = [
        lines[i:i + points_per_slide]
        for i in range(0, len(lines), points_per_slide)
    ]

    slide_number = 1

    image_bytes = ppt_data.get("image_bytes")


    # get slide size
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    for chunk in slides_content:

        slide = prs.slides.add_slide(prs.slide_layouts[1])

        slide_title = generate_slide_title(chunk, content)
        slide.shapes.title.text = slide_title

        tf = slide.placeholders[1].text_frame
        tf.clear()

        for point in chunk:
            p = tf.add_paragraph()
            p.text = point
            p.font.size = Pt(font_size)
            p.level = 0
            p.bullet = True


        # --- ADD IMAGE IF EXISTS ---
        if image_bytes:
            img_stream = io.BytesIO(image_bytes)
            img_width = Inches(5)
            img_height = Inches(3)
            left = (prs.slide_width - img_width) / 2
            top = (prs.slide_height - img_height) / 2
            slide.shapes.add_picture(img_stream, left, top, width=img_width, height=img_height)
            img_stream.close()
            
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(bg_r, bg_g, bg_b)

        add_page_number(slide, slide_number)
        slide_number += 1

    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)

    from werkzeug.utils import secure_filename
    safe_title = secure_filename(title) if title else "presentation"

    if not safe_title.strip():
        safe_title = "presentation"

    return send_file(
        ppt_stream,
        as_attachment=True,
        download_name=f"{safe_title}.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )