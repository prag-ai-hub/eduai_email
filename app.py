import os
import tempfile
from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
import pandas as pd
import smtplib
from email.message import EmailMessage
import threading
import uuid
import time
from flask import jsonify
from mailer import db as mdb
from mailer import custom as mcustom
from mailer import greetings as mgreet
from mailer import utils as mutils
from werkzeug.datastructures import FileStorage
import tempfile
from pathlib import Path
import shutil
from dotenv import load_dotenv
import logging
import re

# try to import premailer for inlining CSS (emails often strip <style> blocks)
try:
    from premailer import transform as premailer_transform
except Exception:
    premailer_transform = None


def sanitize_text_field(s: str) -> str:
    """Remove common labels like 'Hook:' and 'Body:' and trim whitespace."""
    if not s:
        return s
    s = re.sub(r'(?i)\bhook:\s*', '', s)
    s = re.sub(r'(?i)\bbody:\s*', '', s)
    return s.strip()


def inline_css(html: str, keep_style_tags: bool = False) -> str:
    """Inline CSS using premailer when available; otherwise return original HTML.

    keep_style_tags: when True, preserve <style> blocks (useful for preview where we want to show CSS animations). Defaults to False for most sends to maximize inlining support.
    """
    if not html:
        return html
    if premailer_transform:
        try:
            return premailer_transform(html, remove_classes=True, keep_style_tags=keep_style_tags)
        except Exception:
            logging.exception('Premailer transform failed — returning original HTML')
            return html
    return html


def stylize_marketing_body(raw_body: str, product_name: str | None = None) -> str:
    """Convert arbitrary user body text/HTML into a pain-first, stylized HTML fragment suitable for emails.

    - Extracts a short problem statement (first sentence) and highlights it in a styled block.
    - Adds an animated GIF banner (widely supported in Gmail) for subtle motion.
    - Preserves basic links and images in the body if the user included HTML; otherwise converts plain text to paragraphs.
    - Uses inline-friendly styles (no <style> blocks).
    """
    if not raw_body:
        return ''
    raw = raw_body.strip()

    # Remove placeholder labels if present
    raw = sanitize_text_field(raw)

    # Remove leading greetings like "Hi Abhishek," or "Dear Abhishek," if present
    raw = re.sub(r'(?i)^\s*(hi|hello|dear)\s+[A-Z][a-z0-9\-\._ ]{0,40},?\s*\n', '', raw)

    # Remove repeated header lines (company name, tagline) if user pasted whole email header
    raw = re.sub(r"(?im)^(\s*EduAIHub\s*\n\s*Practical AI Tools for Education\s*\n)+", '', raw)
    raw = re.sub(r"(?im)^(\s*EduAIHub\s*\n)+", '', raw)

    # Remove common signature lines (Warm regards, Regards, Thanks, Visit, Unsubscribe, eduaihub) to avoid duplication
    raw = re.sub(r'(?is)\n\s*(warm regards|regards|thanks|thank you)[^\n]*', '', raw)
    raw = re.sub(r'(?is)\n\s*(visit\s+eduaihub[^\n]*|unsubscribe[^\n]*|visit[^\n]*eduaihub[^\n]*)', '', raw)
    raw = re.sub(r'(?is)\n\s*\"[^\n]{0,200}\"\s*—\s*[^\n]{0,100}', '', raw)

    # Trim repeated blank lines
    raw = re.sub(r'\n{3,}', '\n\n', raw).strip()
    # Determine if raw contains HTML tags
    contains_html = bool(re.search(r'<[^>]+>', raw))

    # Extract a short problem statement from the first sentence (plain text)
    plain = re.sub(r'<[^>]+>', '', raw)
    sentences = re.split(r'(?<=[.!?])\s+', plain)
    problem = sentences[0].strip() if sentences and sentences[0].strip() else (plain[:160].strip() + ("..." if len(plain) > 160 else ""))
    if len(problem) > 240:
        problem = problem[:237].rstrip() + '...'

    # Include product context if available
    if product_name:
        problem = f"{problem} — for {product_name}."

    # prefer product image when available, otherwise use neutral animated gif (avoid cats)
    product_img = None
    try:
        for k, v in globals().get('PRODUCTS', {}).items():
            if v.get('name') == product_name and v.get('image_url'):
                product_img = v.get('image_url')
                break
    except Exception:
        product_img = None

    default_gif = os.environ.get('ANIMATED_GIF_URL')
    media_src = product_img or default_gif

    # Build body HTML (preserve user links and images if present)
    if contains_html:
        body_html = raw
    else:
        # Convert plaintext newlines into paragraphs and slightly rephrase for marketing (pain-first)
        paras = [p.strip() for p in raw.split('\n') if p.strip()]
        # If first paragraph doesn't look like a problem, prepend a concise pain-first lead
        lead_keywords = ['grade', 'grading', 'feedback', 'admin', 'administrative', 'time', 'overwhelm', 'overwhelmed', 'drowning', 'tiring', 'burden']
        first_para = paras[0] if paras else ''
        if not any(k in first_para.lower() for k in lead_keywords):
            # craft a pain-first lead using top keywords found in the text if any
            found = [k for k in lead_keywords if any(k in s.lower() for s in paras)]
            if found:
                lead = 'Struggling with ' + ', '.join(found[:2]) + '?' 
            else:
                lead = first_para if first_para else problem
            paras.insert(0, lead)
        paras_html = [f"<p style=\"margin:0 0 12px 0; font-size:14px;\">{p}</p>" for p in paras]
        body_html = '\n'.join(paras_html)

    # Add optional product pains as bullets when product_name is passed (best-effort)
    pains_block = ''
    # find product by name or key
    prod = None
    for k, v in globals().get('PRODUCTS', {}).items():
        if v.get('name') == product_name or k == product_name:
            prod = v
            break
    if prod:
        pains = prod.get('pains', [])
        if pains:
            items = ''.join([f"<li style=\"margin-bottom:6px;color:#4e453f;\">{x}</li>" for x in pains])
            pains_block = f"<ul style=\"margin:8px 0 12px 18px;\">{items}</ul>"

    # CTA: if we have product link, show a small button
    cta_html = ''
    if prod and prod.get('link'):
        # Add a slightly more animated-friendly CTA (box-shadow + gradient) and allow an optional animated gif next to it
        gif = os.environ.get('ANIMATED_GIF_URL')
        gif_html = f"<img src=\"{gif}\" alt=\"\" width=40 style=\"vertical-align:middle;margin-left:8px;border-radius:6px;\" />" if gif else ''
        cta_icon = os.environ.get('CTA_PULSE_URL')
        icon_html = f"<img src=\"{cta_icon}\" alt=\"\" width=18 style=\"vertical-align:middle;margin-left:8px;border-radius:4px;display:inline-block;\" />" if cta_icon else ''
        cta_html = f"<div style=\"margin-top:10px;\"><a href=\"{prod['link']}\" style=\"display:inline-block;padding:10px 14px;background:linear-gradient(90deg,#2fc071,#1aa35a);color:#ffffff;text-decoration:none;border-radius:6px;font-weight:700;box-shadow:0 6px 18px rgba(46,139,87,0.18);\">Learn more about {prod['name']}{icon_html}{gif_html}</a></div>"

    media_html = f"<div style=\"margin-bottom:8px;\"><img src=\"{media_src}\" alt=\"\" width=\"120\" style=\"display:block;border-radius:6px;max-width:100%;height:auto;\" /></div>"

    html = (
        f"<div style=\"padding:0 0 12px 0;\">"
        f"<div style=\"background-color:#fff9d6;border-left:4px solid #f5d84c;padding:12px;border-radius:6px;color:#2f3b1f;font-weight:700;font-size:16px;line-height:1.3;margin-bottom:8px;\">{problem}</div>"
        f"{media_html}"
        f"<div style=\"font-size:14px;color:#234b38;line-height:1.6;\">{body_html}</div>"
        f"{pains_block}"
        f"{cta_html}"
        f"</div>"
    )
    return html


def build_default_structure(fragment_html: str, subject: str = '', sender_name: str = '') -> tuple[str, str]:
    """Return a (structure_html, structure_css) pair for the default email skeleton.

    fragment_html: inner HTML (already wrapped with fragment styles if applicable)
    subject, sender_name: used in header/footer
    """
    # More stylized, animated mobile-friendly skeleton
    structure_css = (
        "body{background-color:#eef2f6;margin:0;padding:0;-webkit-font-smoothing:antialiased;}"
        "table.wrapper{max-width:680px;margin:28px auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e6e9ee;box-shadow:0 6px 30px rgba(37,53,70,0.06);}"
        "td.header{background:linear-gradient(90deg,#2162a6,#2fc071);color:#fff;padding:20px;font-weight:800;font-size:20px;letter-spacing:0.2px}"
        "td.hero{padding:14px;background:linear-gradient(180deg,rgba(255,255,255,0.02),transparent);text-align:center}"
        "td.body{padding:20px;color:#23343a;font-size:15px;line-height:1.7;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif}"
        "td.footer{padding:16px;font-size:12px;color:#7b7b7b;background:#fbfcfd;text-align:center}"
        ".cta{display:inline-block;padding:10px 18px;background:linear-gradient(90deg,#ffb703,#ff7a18);color:#111827;border-radius:8px;text-decoration:none;font-weight:800;margin-top:16px;box-shadow:0 8px 20px rgba(255,122,24,0.18);transition:transform .18s ease,box-shadow .18s ease}"
        ".cta:hover{transform:translateY(-2px);box-shadow:0 12px 28px rgba(34,60,80,0.12)}"
        "@keyframes pulse{0%{transform:scale(1);opacity:1}50%{transform:scale(1.03);opacity:0.9}100%{transform:scale(1);opacity:1}}"
        ".pulse{animation:pulse 3s infinite ease-in-out}"
        ".divider{height:6px;background:linear-gradient(90deg,#1aa35a,#2fc071);border-radius:8px;margin:12px 0}"
        ".ai-fragment h1,.ai-fragment h2{animation:fadeInDown .6s ease both}"
        "@keyframes fadeInDown{from{opacity:0;transform:translateY(-4px)}to{opacity:1;transform:none}}"
    )
    # header shows subject/title and optional sender; include an optional animated GIF if configured
    gif = os.environ.get('ANIMATED_GIF_URL')
    gif_html = f"<img src=\"{gif}\" alt=\"\" width=76 style=\"vertical-align:middle;border-radius:8px;margin-left:12px;\"/>" if gif else ''
    # Use inline styling fallbacks for higher email-client compatibility
    header_inline = "background:linear-gradient(90deg,#ff7ab6,#4cc9f0);color:#ffffff;padding:20px;font-weight:900;font-size:20px;"
    header_html = f"<td class=\"header\" style=\"{header_inline}\"><div style=\"display:flex;align-items:center;justify-content:space-between;gap:12px;\"><div>{subject or 'Update from EduAI'}</div><div>{gif_html}</div></div></td>"
    footer_html = f"<td class=\"footer\" style=\"padding:16px;font-size:12px;color:#6b7280;background:#f8fbff;text-align:center\">{sender_name or 'EduAI Hub'} • <a href=\"https://eduaihub.in\">eduaihub.in</a></td>"
    # include a subtle divider (inline) and leave the CTA up to the fragment, but add a placeholder wrapper
    divider_inline = "height:8px;background:linear-gradient(90deg,#ff7ab6,#ffd166);border-radius:12px;margin:14px 0;display:block"
    structure_html = (
        "<table class=\"wrapper\" cellpadding=0 cellspacing=0 width=100%>"
        f"<tr>{header_html}</tr>"
        f"<tr><td class=\"hero\"><div class=\"divider pulse\" style=\"{divider_inline}\"></div></td></tr>"
        f"<tr><td class=\"body\" style=\"padding:22px;color:#102a43;font-size:15px;line-height:1.75;background:linear-gradient(180deg,#ffffff,#fbfdff);\">{fragment_html}</td></tr>"
        f"<tr>{footer_html}</tr>"
        "</table>"
    )
    return structure_html, structure_css


def normalize_fragment_html(content: str) -> str:
    """Ensure fragment content is wrapped into block elements with inline fallbacks so styles apply.

    - If the content already contains block-level HTML, wrap it in a container `div.ai-body-content`.
    - Otherwise, split on blank lines and convert to <p> with inline paragraph styles.
    """
    if not content:
        return ''
    # If already has block-level elements, just wrap
    if re.search(r'<\s*(p|div|ul|ol|table|h[1-6]|blockquote)\b', content, re.I):
        return f"<div class=\"ai-body-content\">{content}</div>"

    # Plain text: split on double newlines into paragraphs
    paras = [p.strip() for p in re.split(r'\n{2,}|\r\n{2,}', content) if p.strip()]
    if not paras:
        paras = [content.strip()]
    p_style = 'margin:0 0 12px 0;color:#1f3a5f;font-size:15px;line-height:1.7'
    paras_html = ''.join([f"<p style=\"{p_style}\">" + p.replace('\n','<br/>') + "</p>" for p in paras])
    return f"<div class=\"ai-body-content\">{paras_html}</div>"


def has_meaningful_body(html_fragment: str, min_chars: int = 30, min_words: int = 5) -> bool:
    """Return True if the given HTML or fragment likely contains a meaningful message body.

    This strips HTML tags, removes common header/footer lines we add automatically, collapses
    whitespace and checks simple length / word heuristics. Used to prevent accidental
    sends of header/footer-only content.
    """
    if not html_fragment:
        return False
    # strip tags
    text = re.sub(r'<[^>]+>', '', html_fragment)
    # remove common header/footer lines users sometimes paste
    text = re.sub(r'(?i)eduai\s*hub\s*(•|-)\s*eduaihub\.in\s*(•|-)\s*unsubscribe', '', text)
    text = re.sub(r'(?i)unsubscribe', '', text)
    text = re.sub(r'(?i)visit\s+eduaihub', '', text)
    text = re.sub(r'(?i)where education meets intelligence', '', text)
    # strip common greetings alone (hi, hello, dear, regards)
    text = re.sub(r'(?i)^\s*(hi|hello|dear|regards|thanks|thank you)\s*[,.]?\s*$', '', text)
    text = text.strip()
    text = re.sub(r'\s+', ' ', text)
    if len(text) < min_chars:
        return False
    if len(text.split()) < min_words:
        return False
    if not re.search(r'[A-Za-z0-9]', text):
        return False
    return True


logging.basicConfig(level=logging.INFO)

load_dotenv()

PREMAILER_AVAILABLE = premailer_transform is not None

# Detect OpenAI availability (set by .env via load_dotenv above or environment)
try:
    import openai as _openai_check
    OPENAI_AVAILABLE = bool(os.environ.get('OPENAI_API_KEY') or getattr(_openai_check, 'api_key', None))
except Exception:
    OPENAI_AVAILABLE = False

SMTP_SERVER = os.environ.get('SMTP_SERVER')
SMTP_PORT = int(os.environ.get('SMTP_PORT', 587))
SMTP_EMAIL = os.environ.get('SMTP_EMAIL')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD')
SMTP_DEBUG = os.environ.get('SMTP_DEBUG', '0') in ('1', 'true', 'True')

# In-memory task tracking
tasks = {}
tasks_lock = threading.Lock()

ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'csv'}

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'devsecret')
# expose OpenAI availability to templates
app.config['OPENAI_AVAILABLE'] = OPENAI_AVAILABLE


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_emails_from_dataframe(df):
    for col in df.columns:
        if 'email' in str(col).lower():
            return df[col].dropna().astype(str).str.strip().unique().tolist()
    first = df.columns[0]
    return df[first].dropna().astype(str).str.strip().unique().tolist()


def send_bulk_emails(recipients, subject, html_body, sender_display=None):
    results = {'sent': 0, 'failed': []}
    if not SMTP_SERVER or not SMTP_EMAIL or not SMTP_PASSWORD:
        raise RuntimeError('SMTP settings not configured in environment')

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.ehlo()
        if SMTP_DEBUG:
            smtp.set_debuglevel(1)
        smtp.starttls()
        smtp.login(SMTP_EMAIL, SMTP_PASSWORD)

        for r in recipients:
            try:
                msg = EmailMessage()
                display_from = f"{sender_display} <{SMTP_EMAIL}>" if sender_display else SMTP_EMAIL
                msg['From'] = display_from
                msg['To'] = r
                msg['Subject'] = subject
                msg.set_content('This email requires an HTML-capable client.')
                msg.add_alternative(html_body, subtype='html')
                smtp.send_message(msg)
                results['sent'] += 1
            except Exception as e:
                logging.exception('Failed to send to %s', r)
                results['failed'].append({'email': r, 'error': str(e)})
    return results


def send_task(task_id, recipients, subject, html_body, sender_display=None):
    """Background task that sends emails and updates the tasks dict."""
    try:
        with tasks_lock:
            tasks[task_id]['status'] = 'running'
            tasks[task_id]['total'] = len(recipients)
            tasks[task_id]['sent'] = 0
            tasks[task_id]['failed'] = 0

        if not SMTP_SERVER or not SMTP_EMAIL or not SMTP_PASSWORD:
            raise RuntimeError('SMTP settings not configured')

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=60) as smtp:
            smtp.ehlo()
            if SMTP_DEBUG:
                smtp.set_debuglevel(1)
            smtp.starttls()
            smtp.login(SMTP_EMAIL, SMTP_PASSWORD)

            for r in recipients:
                try:
                    msg = EmailMessage()
                    display_from = f"{sender_display} <{SMTP_EMAIL}>" if sender_display else SMTP_EMAIL
                    msg['From'] = display_from
                    msg['To'] = r
                    msg['Subject'] = subject
                    msg.set_content('This email requires an HTML-capable client.')
                    msg.add_alternative(html_body, subtype='html')
                    smtp.send_message(msg)
                    with tasks_lock:
                        tasks[task_id]['sent'] += 1
                except Exception as e:
                    logging.exception('Error sending to %s', r)
                    with tasks_lock:
                        tasks[task_id]['failed'] += 1
                time.sleep(0.1)

        with tasks_lock:
            tasks[task_id]['status'] = 'done'
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        logging.exception('Background send task failed (task_id=%s)', task_id)
        with tasks_lock:
            tasks[task_id]['status'] = 'error'
            tasks[task_id]['error'] = tb


@app.route('/test-smtp', methods=['GET'])
def test_smtp():
    """Quick endpoint to test SMTP connectivity and login without sending messages."""
    try:
        if not SMTP_SERVER or not SMTP_EMAIL or not SMTP_PASSWORD:
            return {'ok': False, 'error': 'SMTP environment variables missing'}, 400
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=15) as smtp:
            smtp.ehlo()
            if SMTP_DEBUG:
                smtp.set_debuglevel(1)
            smtp.starttls()
            smtp.login(SMTP_EMAIL, SMTP_PASSWORD)
        return {'ok': True, 'message': 'SMTP login successful'}
    except Exception as e:
        logging.exception('SMTP test failed')
        return {'ok': False, 'error': str(e)}, 500


@app.route('/')
def dashboard():
    return render_template('dashboard.html')


@app.route('/custom', methods=['GET'])
def custom_form():
    return render_template('custom_send.html', products=PRODUCTS)


@app.route('/custom/send', methods=['POST'])
def custom_send():
    # Accept optional recipients file or emails list, save attachments and generate content, then show preview for confirmation
    file = request.files.get('file')
    emails_raw = request.form.get('emails') or ''
    name = request.form.get('name') or ''
    product_key = request.form.get('product_key') or ''
    description = sanitize_text_field(request.form.get('description') or '')
    files = request.files.getlist('files')
    ai_personalize = request.form.get('ai_personalize') == '1'

    # Template selection and optional event details (defaults so they are always defined)
    template_name = request.form.get('template') or ''
    cta_text = request.form.get('cta_text') or ''
    cta_link = request.form.get('cta_link') or ''
    event_date = request.form.get('event_date') or ''
    event_time = request.form.get('event_time') or ''
    event_location = request.form.get('event_location') or ''

    # Optional structure-only skeleton and user CSS
    structure_html = request.form.get('structure_html') or ''
    structure_css = request.form.get('structure_css') or ''

    # Early validation: require meaningful description content before processing
    if not description or not description.strip():

        flash('Please provide a message in the Description field.', 'warning')
        return redirect(request.url)
    # Check that description is not just a greeting
    desc_check = re.sub(r'(?i)^\s*(hi|hello|dear|regards|thanks|thank you)\s*[,.]?\s*$', '', description.strip())
    if not desc_check or len(desc_check.split()) < 5:
        flash('Message is too short. Please provide at least 5 words of content (greetings alone are not sufficient).', 'warning')
        return redirect(request.url)

    recipients = []
    tmpdir = None
    try:
        if file and file.filename:
            tmpdir = tempfile.mkdtemp()
            path = os.path.join(tmpdir, secure_filename(file.filename))
            file.save(path)
            try:
                if path.lower().endswith('.csv'):
                    df = pd.read_csv(path)
                else:
                    df = pd.read_excel(path, engine='openpyxl')
            except Exception as e:
                flash('Failed to read recipients file: ' + str(e), 'danger')
                return redirect(request.url)

            emails = extract_emails_from_dataframe(df)
            name_col = None
            for c in df.columns:
                if 'name' in str(c).lower() or 'first' in str(c).lower() or 'full' in str(c).lower():
                    name_col = c
                    break
            if name_col:
                for e in emails:
                    matches = df[df.apply(lambda r: any(str(v).lower().find(str(e).lower()) != -1 for v in r.values), axis=1)]
                    name_val = ''
                    if not matches.empty and name_col in matches.columns:
                        try:
                            name_val = str(matches.iloc[0][name_col])
                        except Exception:
                            name_val = ''
                    recipients.append({'email': e, 'name': name_val})
            else:
                recipients = [{'email': e, 'name': ''} for e in emails]
        else:
            recipients = [{'email': e.strip(), 'name': name} for e in emails_raw.splitlines() if e.strip()]

        if not recipients:
            flash('Please provide recipients via upload or list', 'warning')
            return redirect(request.url)

        product_pains = None
        product_name = None
        if product_key:
            product = PRODUCTS.get(product_key)
            if product:
                product_pains = product.get('pains')
                product_name = product.get('name')

        # Defaults: structure-only OpenAI formatting and personalization are applied by default (no checkboxes in UI).
        use_ai_structure = True
        # use_ai flag kept for compatibility with downstream handlers; default it to False since we always use structure-only logic
        use_ai = False

        # Prepare variables used by AI formatting
        html_body = ''  # full HTML if AI returns a complete document
        body_fragment = ''  # fallback fragment content

        # Run structure-only formatting (preserve wording, insert [[RECIPIENT_NAME]] placeholder in greetings)
        first = recipients[0]
        try:
            # Prefer generating a full, stylized HTML email from OpenAI when available
            if OPENAI_AVAILABLE:
                try:
                    subj_ai, full_html = mcustom.generate_full_html(description, subject=subj, recipient_name='', product_name=product_name, product_pains=product_pains)
                    # If model provided a subject suggestion, adopt it unless user overrode
                    user_subj = request.form.get('subject')
                    subj = user_subj if user_subj else (subj_ai or subj)
                    # Keep full_html as the preview email body
                    body_fragment = ''
                    # For preview, substitute the first recipient's name into the placeholder so the user sees a realistic preview
                    try:
                        preview_html = full_html.replace('[[RECIPIENT_NAME]]', first.get('name',''))
                    except Exception:
                        preview_html = full_html
                    html_body = preview_html
                    # remember the canonical AI html for sending (contains [[RECIPIENT_NAME]] placeholder)
                    canonical_ai_html = full_html
                except Exception:
                    # fallback to structure-only formatting when full HTML generation fails
                    subj_ret, rewritten = mcustom.rewrite_body(description, recipient_name='', product_name=product_name, product_pains=product_pains, structure_only=True)
                    user_subj = request.form.get('subject')
                    subj = user_subj if user_subj else (subj_ret or subj)
                    body_fragment = rewritten
            else:
                subj_ret, rewritten = mcustom.rewrite_body(description, recipient_name='', product_name=product_name, product_pains=product_pains, structure_only=True)
                user_subj = request.form.get('subject')
                subj = user_subj if user_subj else (subj_ret or subj)
                body_fragment = rewritten
        except Exception:
            subj = request.form.get('subject') or f'Update from EduAI Hub'
            body_fragment = stylize_marketing_body(description, product_name=product_name)

        # Determine whether the fragment contains a greeting (placeholder or explicit greeting)
        fragment_has_greeting = '[[RECIPIENT_NAME]]' in (body_fragment or '') or bool(re.search(r'(?im)^(\s*(hi|hello|dear)\s+)', body_fragment or ''))
        # If not, we'll need to inject a personalized greeting per recipient when sending
        fragment_needs_greeting = not fragment_has_greeting

        # For preview, show exactly what will be sent to the first recipient: inject greeting if needed
        first_name = first.get('name','')
        p_style = "margin:0 0 12px 0; font-size:15px; color:#234b38; line-height:1.6; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif"
        fragment_style = "<style>.ai-fragment{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif;font-size:15px;color:#1f3a5f;line-height:1.7;background:linear-gradient(180deg,#ffffff,#fbfdff);padding:10px;border-radius:10px}.ai-fragment p{margin:0 0 12px 0}.ai-fragment .lead{font-weight:900;color:#0b60a6;margin-bottom:8px;font-size:18px}.ai-fragment .cta-inline{display:inline-block;padding:10px 14px;background:linear-gradient(90deg,#ff7ab6,#6bdeff);color:#05233a;border-radius:10px;text-decoration:none;font-weight:900;box-shadow:0 10px 28px rgba(107,222,255,0.12);transition:transform .18s}.ai-fragment .cta-inline:hover{transform:translateY(-3px)}.ai-fragment .badge{display:inline-block;background:#ffd166;color:#6b3b00;padding:6px 8px;border-radius:999px;font-weight:900;margin-right:8px}.ai-decor{display:flex;justify-content:flex-end;gap:8px;margin-bottom:8px}.spark{display:inline-block;width:10px;height:10px;border-radius:50%;background:linear-gradient(90deg,#ff7ab6,#ffd166);box-shadow:0 8px 20px rgba(255,122,182,0.12);animation:confetti 3s linear infinite}@keyframes confetti{0%{transform:translateY(-4px) rotate(0);opacity:1}50%{transform:translateY(2px) rotate(180deg);opacity:0.8}100%{transform:translateY(-2px) rotate(360deg);opacity:0.3}}@keyframes fragBounce{0%{transform:translateY(-6px);opacity:0}60%{transform:translateY(3px);opacity:1}100%{transform:none}}.ai-fragment{animation:fragBounce .6s cubic-bezier(.17,.67,.3,1) both}.ai-body-content p{margin:0 0 12px 0;color:#102a43;font-size:15px;line-height:1.7}.ai-body-content h2{color:#ff7ab6;margin:0 0 8px 0}.ai-body-content ul li{margin:6px 0;padding-left:6px;color:#1f3a5f}.ai-body-content blockquote{border-left:4px solid #ffd166;padding:8px 12px;background:#fffaf0;color:#6b4a00;border-radius:6px}.ai-fragment .spark{width:10px;height:10px}@keyframes fragFadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}.ai-body-content p{animation:fragFadeIn .5s ease both}</style>"
        def _wrap_fragment(content):
            decor_html = "<div class='ai-decor'><span class='spark'></span><span class='spark'></span><span class='spark'></span></div>"
            processed = normalize_fragment_html(content)
            return fragment_style + f"<div class=\"ai-fragment\" style=\"padding:10px;border-radius:10px;background:linear-gradient(180deg,#ffffff,#fbfdff);\">{decor_html}{processed}</div>"

        if fragment_has_greeting:
            frag = (body_fragment or '').replace('[[RECIPIENT_NAME]]', first_name)
        else:
            greeting_html = f"<p style=\"{p_style}\">Hi {first_name or 'Educator'},</p>"
            frag = (greeting_html + (body_fragment or '')).replace('[[RECIPIENT_NAME]]', first_name)
        fragment_preview = _wrap_fragment(frag)

        # Choose template: default or themed event template based on user selection
        if 'html_body' not in locals() or not html_body:
            template_name = request.form.get('template') or ''
            # If user provided a custom structure skeleton, inject the AI fragment into it and include any user CSS
            if structure_html:
                # insert fragment_preview into the <td class="body"> ... </td> block if present; otherwise append
                try:
                    if re.search(r'(?is)<td[^>]*class=["\']body["\'][^>]*>.*?<\/td>', structure_html):
                        # replace inner content
                        html_with_body = re.sub(r'(?is)(<td[^>]*class=["\']body["\'][^>]*>).*?(<\/td>)', r"\1" + fragment_preview + r"\2", structure_html)
                    elif '[[AI_BODY]]' in structure_html:
                        html_with_body = structure_html.replace('[[AI_BODY]]', fragment_preview)
                    else:
                        html_with_body = structure_html.replace('</table>', f'<tr><td class="body">{fragment_preview}</td></tr></table>', 1)
                except Exception:
                    html_with_body = structure_html + fragment_preview

                # wrap with head including any user CSS
                head_css = f"<style>{structure_css}</style>" if structure_css else ''
                html_body = f"<html><head>{head_css}</head><body>{html_with_body}</body></html>"
            elif template_name == 'event':
                # Gather optional event fields from the form (may be empty)
                event_date = request.form.get('event_date') or ''
                event_time = request.form.get('event_time') or ''
                event_location = request.form.get('event_location') or 'LIVE ON ZOOM'
                cta_text = request.form.get('cta_text') or 'Register for tomorrow\'s live session'
                cta_link = request.form.get('cta_link') or '#'
                html_body = render_template('email_template_theme.html', recipient_name=first_name, ai_body=fragment_preview, subject=subj, title=subj, event_date=event_date, event_time=event_time, event_location=event_location, cta_text=cta_text, cta_link=cta_link, sender_name=name)
            else:
                # If user did not provide a custom structure, build a default skeleton and include fragment styles/CSS
                if not structure_html:
                    generated_structure_html, generated_structure_css = build_default_structure(fragment_preview, subject=subj, sender_name=name)
                    # prepare preview HTML with CSS in head so the browser shows styling
                    html_body = f"<html><head><style>{generated_structure_css}</style></head><body>{generated_structure_html}</body></html>"
                    # pass generated structure back so send task can reuse it
                    structure_html = generated_structure_html
                    structure_css = generated_structure_css
                else:
                    html_body = render_template('email_template_custom.html', recipient_name='', ai_hook='', ai_body=fragment_preview, suppress_greeting=True)

        if PREMAILER_AVAILABLE and html_body:
            # Preserve <style> blocks for the browser preview so animations and advanced styles are visible to the user.
            html_body = inline_css(html_body, keep_style_tags=True)
        else:
            flash('Premailer not installed — styles may not appear in some email clients. Install with `pip install premailer`.', 'warning')

        # save attachments temporarily and pass their paths to preview confirm step
        tempdir_att = tempfile.mkdtemp()
        saved = []
        for f in files:
            if isinstance(f, FileStorage) and f.filename:
                path = os.path.join(tempdir_att, secure_filename(f.filename))
                f.save(path)
                saved.append(path)

        recipient_data = '\n'.join([f"{r['email']}||{r['name']}" for r in recipients])

        return render_template('preview.html', emails=[r['email'] for r in recipients], count=len(recipients), subject=subj, sender_name=name, email_html=html_body, custom_attachments='||'.join(saved), custom_mode='custom', recipient_data=recipient_data, ai_personalize='1' if ai_personalize else '0', email_fragment=body_fragment, use_ai='1' if use_ai else '0', template=template_name, cta_text=cta_text if template_name=='event' else '', cta_link=cta_link if template_name=='event' else '', event_date=event_date if template_name=='event' else '', event_time=event_time if template_name=='event' else '', event_location=event_location if template_name=='event' else '', structure_html=structure_html, structure_css=structure_css)
    finally:
        # keep tempdirs for confirm step; cleanup after sending
        pass


@app.route('/custom-start-send', methods=['POST'])
def custom_start_send():
    # Accept recipient_data (email||name per line) and attachments list
    recipient_data_raw = request.form.get('recipient_data') or ''
    subject = request.form.get('subject') or 'Update from EduAI'
    sender_name = request.form.get('sender_name') or ''
    attachments_raw = request.form.get('custom_attachments') or ''
    attachments = [p for p in attachments_raw.split('||') if p]
    email_html = request.form.get('email_html') or ''
    ai_personalize = request.form.get('ai_personalize') == '1'
    use_ai = request.form.get('use_ai') == '1'
    product_key = request.form.get('product_key') or ''
    description = request.form.get('description') or ''
    email_fragment = request.form.get('email_fragment') or ''

    recipients = []
    for line in recipient_data_raw.splitlines():
        if not line.strip():
            continue
        parts = line.split('||')
        email = parts[0].strip()
        name = parts[1].strip() if len(parts) > 1 else ''
        recipients.append({'email': email, 'name': name})

    if not recipients:
        flash('No recipients to send to', 'warning')
        return redirect(url_for('custom_form'))

    # Prevent accidental sends when the formatted message lacks a real body
    # (e.g., only contains header/footer or is empty). Require the user to
    # edit/confirm the message in the preview first.
    if email_fragment:
        if not has_meaningful_body(email_fragment):
            logging.warning('Blocked send: fragment rejected by validation. Fragment: %s', email_fragment[:300])
            flash('Message appears to be empty or contains only a greeting/header/footer. Please add more content to your message (at least 5 words of body text).', 'warning')
            return redirect(url_for('custom_form'))
    elif email_html:
        if not has_meaningful_body(email_html):
            logging.warning('Blocked send: preview HTML rejected by validation. HTML snippet: %s', email_html[:300])
            flash('Preview looks empty or contains only header/footer. Please review the preview and add more content before sending.', 'warning')
            return redirect(url_for('preview'))
    else:
        flash('No message content detected. Please provide a message in the editor.', 'warning')
        return redirect(url_for('custom_form'))

    if request.form.get('dry_run') == '1':
        flash('Dry run: no emails were sent. Preview only.', 'info')
        results = {'sent': 0, 'failed': []}
        return render_template('result.html', results=results)

    # Start background send task
    task_id = str(uuid.uuid4())
    with tasks_lock:
        tasks[task_id] = {'status': 'pending', 'total': len(recipients), 'sent': 0, 'failed': 0}

    # Resolve product context for personalization
    product_pains = None
    product_name = None
    if product_key:
        p = PRODUCTS.get(product_key)
        if p:
            product_pains = p.get('pains')
            product_name = p.get('name')

    # Read fragment_has_greeting and pass structure flag to the background send task
    fragment_has_greeting = request.form.get('fragment_has_greeting') == '1'
    use_ai_structure = True
    template_name = request.form.get('template') or ''
    cta_text = request.form.get('cta_text') or ''
    cta_link = request.form.get('cta_link') or ''
    event_date = request.form.get('event_date') or ''
    event_time = request.form.get('event_time') or ''
    event_location = request.form.get('event_location') or ''

    # read structure HTML/CSS passed from preview (may be empty)
    structure_html = request.form.get('structure_html') or ''
    structure_css = request.form.get('structure_css') or ''

    thread = threading.Thread(target=send_custom_task, args=(task_id, recipients, subject, sender_name, attachments, email_html, ai_personalize, use_ai, use_ai_structure, description, email_fragment, fragment_has_greeting, product_key, product_name, product_pains, template_name, cta_text, cta_link, event_date, event_time, event_location, structure_html, structure_css), daemon=True)
    thread.start()
    flash('Customized emails queued for sending (check progress).', 'success')
    return redirect(url_for('progress', task_id=task_id))


@app.route('/send-preview', methods=['POST'])
def send_preview():
    """Send the exact preview HTML to a single test email address for verification."""
    test_email = request.form.get('test_email') or ''
    subj = request.form.get('subject') or 'Preview — EduAI'
    sender_name = request.form.get('sender_name') or ''
    email_html = request.form.get('email_html') or ''

    if not test_email:
        flash('Please provide a destination email address to send the preview to.', 'warning')
        return redirect(request.referrer or url_for('dashboard'))

    # Use the same send helper so behavior matches the normal send path
    try:
        # inline CSS for sending to ensure it matches preview rendering in most clients
        if PREMAILER_AVAILABLE:
            send_html = inline_css(email_html, keep_style_tags=True)
        else:
            send_html = email_html

        results = send_bulk_emails([test_email], subj, send_html, sender_display=sender_name)
        if results.get('failed'):
            flash(f"Preview send completed but some sends failed: {results['failed']}", 'warning')
        else:
            flash('Preview sent successfully — check the inbox of the test address.', 'success')
    except Exception as e:
        logging.exception('Failed to send preview')
        flash('Failed to send preview: ' + str(e), 'danger')

    return redirect(request.referrer or url_for('dashboard'))


def send_custom_task(task_id, recipients, subject, sender_display=None, attachments=None, body_html=None, ai_personalize=False, use_ai=False, use_ai_structure=False, description='', body_fragment='', fragment_has_greeting=False, product_key=None, product_name=None, product_pains=None, template_name='', cta_text='', cta_link='', event_date='', event_time='', event_location='', structure_html='', structure_css=''):
    """Send emails using the EXACT preview HTML with per-recipient name personalization.
    
    The preview HTML (body_html) is the authoritative template. We personalize it per recipient
    by replacing a special marker with each recipient's name. This ensures what you see in
    preview is exactly what gets sent (with names personalized).
    """
    try:
        with tasks_lock:
            tasks[task_id]['status'] = 'running'
            tasks[task_id]['total'] = len(recipients)
            tasks[task_id]['sent'] = 0
            tasks[task_id]['failed'] = 0

        # Use the preview HTML as the base template. The preview was generated for the first recipient,
        # so we need to extract a generic version by replacing the first recipient's name with a marker.
        first_recipient_name = recipients[0].get('name', '') if recipients else ''
        
        # Create a generic template from the original fragment so we guarantee a placeholder
        p_style = "margin:0 0 12px 0; font-size:15px; color:#234b38; line-height:1.6; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif"
        greeting_marker = "___RECIPIENT_NAME_PLACEHOLDER___"

        # Prefer building the generic template from `body_fragment` (authoritative fragment used for preview)
        if body_fragment:
            # If user provided a structure_html skeleton, build a generic HTML from it that contains the greeting marker
            if structure_html:
                # Prepare a frag_for_send that uses the internal greeting marker
                if '[[RECIPIENT_NAME]]' in body_fragment:
                    frag_for_send = body_fragment.replace('[[RECIPIENT_NAME]]', greeting_marker)
                elif not fragment_has_greeting and first_recipient_name:
                    frag_for_send = f"<p style=\"{p_style}\">Hi {greeting_marker},</p>" + body_fragment
                else:
                    frag_for_send = body_fragment
                # apply consistent fragment styles so content keeps its look when inserted into skeleton
                fragment_style = "<style>.ai-fragment{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif;font-size:15px;color:#1f3a5f;line-height:1.7;background:linear-gradient(180deg,#ffffff,#fbfdff);padding:6px;border-radius:8px}.ai-fragment p{margin:0 0 12px 0}.ai-fragment .lead{font-weight:900;color:#0b60a6;margin-bottom:8px;font-size:17px}.ai-fragment .cta-inline{display:inline-block;padding:10px 14px;background:linear-gradient(90deg,#ff7ab6,#6bdeff);color:#05233a;border-radius:10px;text-decoration:none;font-weight:900;box-shadow:0 10px 28px rgba(107,222,255,0.12);transition:transform .18s}.ai-fragment .cta-inline:hover{transform:translateY(-3px)}.ai-fragment .badge{display:inline-block;background:#ffd166;color:#6b3b00;padding:6px 8px;border-radius:999px;font-weight:900;margin-right:8px}.ai-decor{display:flex;justify-content:flex-end;gap:8px;margin-bottom:8px}.spark{display:inline-block;width:10px;height:10px;border-radius:50%;background:linear-gradient(90deg,#ff7ab6,#ffd166);box-shadow:0 8px 20px rgba(255,122,182,0.12);animation:confetti 3s linear infinite}@keyframes confetti{0%{transform:translateY(-4px) rotate(0);opacity:1}50%{transform:translateY(2px) rotate(180deg);opacity:0.8}100%{transform:translateY(-2px) rotate(360deg);opacity:0.3}}@keyframes fragFade{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}</style>"
                decor_html = "<div class='ai-decor'><span class='spark'></span><span class='spark'></span><span class='spark'></span></div>"
                processed_frag = normalize_fragment_html(frag_for_send)
                frag_for_send = fragment_style + f"<div class=\"ai-fragment\" style=\"padding:10px;border-radius:10px;background:linear-gradient(180deg,#ffffff,#fbfdff);\">{decor_html}{processed_frag}</div>"
                # Inject fragment into structure_html
                try:
                    if re.search(r'(?is)<td[^>]*class=["\']body["\'][^>]*>.*?<\/td>', structure_html):
                        generic_html = re.sub(r'(?is)(<td[^>]*class=["\']body["\'][^>]*>).*?(<\/td>)', r"\1" + frag_for_send + r"\2", structure_html)
                    elif '[[AI_BODY]]' in structure_html:
                        generic_html = structure_html.replace('[[AI_BODY]]', frag_for_send)
                    else:
                        generic_html = structure_html.replace('</table>', f'<tr><td class="body">{frag_for_send}</td></tr></table>', 1)
                except Exception:
                    generic_html = structure_html + frag_for_send
                # include any user CSS
                if structure_css:
                    generic_html = f"<html><head><style>{structure_css}</style></head><body>{generic_html}</body></html>"
                if PREMAILER_AVAILABLE:
                    generic_html = inline_css(generic_html, keep_style_tags=False)
                suppress = True
            else:
                # If the fragment contains the placeholder token, replace it with our internal marker
                if '[[RECIPIENT_NAME]]' in body_fragment:
                    frag_for_send = body_fragment.replace('[[RECIPIENT_NAME]]', greeting_marker)
                    suppress = True
                elif not fragment_has_greeting and first_recipient_name:
                    # Inject a greeting that uses the marker (keeps same p_style used for preview)
                    frag_for_send = f"<p style=\"{p_style}\">Hi {greeting_marker},</p>" + body_fragment
                    suppress = True
                else:
                    frag_for_send = body_fragment
                    suppress = fragment_has_greeting
                # apply consistent fragment styles so content keeps its look when rendered by the template
                fragment_style = "<style>.ai-fragment{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif;font-size:15px;color:#1f3a5f;line-height:1.7;background:linear-gradient(180deg,#ffffff,#fbfdff);padding:10px;border-radius:10px}.ai-fragment p{margin:0 0 12px 0}.ai-fragment .lead{font-weight:900;color:#0b60a6;margin-bottom:8px;font-size:17px}.ai-decor{display:flex;justify-content:flex-end;gap:8px;margin-bottom:8px}.spark{display:inline-block;width:10px;height:10px;border-radius:50%;background:linear-gradient(90deg,#ff7ab6,#ffd166);box-shadow:0 8px 20px rgba(255,122,182,0.12);animation:confetti 3s linear infinite}@keyframes confetti{0%{transform:translateY(-4px) rotate(0);opacity:1}50%{transform:translateY(2px) rotate(180deg);opacity:0.8}100%{transform:translateY(-2px) rotate(360deg);opacity:0.3}}@keyframes fragBounce{0%{transform:translateY(-6px);opacity:0}60%{transform:translateY(3px);opacity:1}100%{transform:none}}.ai-fragment{animation:fragBounce .6s cubic-bezier(.17,.67,.3,1) both}</style>"
                decor_html = "<div class='ai-decor'><span class='spark'></span><span class='spark'></span><span class='spark'></span></div>"
                processed_frag = normalize_fragment_html(frag_for_send)
                frag_for_send = fragment_style + f"<div class=\"ai-fragment\" style=\"padding:10px;border-radius:10px;background:linear-gradient(180deg,#ffffff,#fbfdff);\">{decor_html}{processed_frag}</div>"
                # Render the full template from this fragment and fully inline CSS for sending
                with app.app_context():
                    if template_name == 'event':
                        generic_html = render_template('email_template_theme.html', recipient_name='', ai_body=frag_for_send, subject=subject, title=subject, event_date=event_date, event_time=event_time, event_location=event_location, cta_text=cta_text or 'Register', cta_link=cta_link or '#', sender_name=sender_display, footer_text='Solving complex business problems with intelligent automation solutions')
                    elif not structure_html:
                        # No user skeleton provided — build default skeleton on the backend and inline CSS for sending
                        generated_structure_html, generated_structure_css = build_default_structure(frag_for_send, subject=subject, sender_name=sender_display)
                        generic_html = f"<html><head><style>{generated_structure_css}</style></head><body>{generated_structure_html}</body></html>"
                    else:
                        generic_html = render_template('email_template_custom.html', recipient_name='', ai_hook='', ai_body=frag_for_send, suppress_greeting=suppress)
                    if PREMAILER_AVAILABLE:
                        generic_html = inline_css(generic_html, keep_style_tags=False)
        elif body_html:
            # Fallback: try to produce a generic version by replacing the first recipient's name if present
            generic_html = body_html
            if first_recipient_name:
                generic_html = generic_html.replace(f"Hi {first_recipient_name},", f"Hi {greeting_marker},")
                generic_html = generic_html.replace(f"Dear {first_recipient_name},", f"Dear {greeting_marker},")
                generic_html = generic_html.replace(f"Hello {first_recipient_name},", f"Hello {greeting_marker},")
                generic_html = generic_html.replace(f"Hi&nbsp;{first_recipient_name},", f"Hi&nbsp;{greeting_marker},")
                generic_html = generic_html.replace(f"Dear&nbsp;{first_recipient_name},", f"Dear&nbsp;{greeting_marker},")
            # ensure CSS inlined for sending
            if PREMAILER_AVAILABLE:
                generic_html = inline_css(generic_html, keep_style_tags=False)
        else:
            generic_html = ''

        for r in recipients:
            email = r.get('email')
            name = r.get('name') or 'Educator'
            try:
                # Ensure any [[RECIPIENT_NAME]] placeholders are converted to our internal marker
                generic_html = generic_html.replace('[[RECIPIENT_NAME]]', greeting_marker)
                # Personalize the generic HTML for this recipient
                final_html = generic_html.replace(greeting_marker, name)
                use_subj = subject
                
                logging.info('Sending preview-based email to %s with personalized greeting for %s', email, name)

                # Defensive check: ensure final_html contains meaningful body before sending
                if not has_meaningful_body(final_html):
                    logging.warning('Skipping send to %s due to empty or header-only body', email)
                    with tasks_lock:
                        tasks[task_id]['failed'] += 1
                    try:
                        mdb.log_entry(email, 'custom', use_subj, 'skipped-empty-body')
                    except Exception:
                        logging.exception('Failed to log skipped empty body')
                    continue

                # Log a short debug snippet of the HTML actually being sent to help diagnose missing-body issues
                try:
                    logging.debug('Sending to %s: subject=%s, html_len=%s, html_snippet=%s', email, use_subj, len(final_html) if final_html else 0, (final_html or '')[:300].replace('\n', ' '))
                except Exception:
                    pass
                ok, err = mutils.send_email_with_attachments(email, use_subj, final_html, attachments=attachments or None, sender_name=sender_display)
                if ok:
                    with tasks_lock:
                        tasks[task_id]['sent'] += 1
                    mdb.log_entry(email, 'custom', use_subj, 'sent')
                else:
                    with tasks_lock:
                        tasks[task_id]['failed'] += 1
                    mdb.log_entry(email, 'custom', use_subj, f'failed: {err}')
            except Exception as e:
                with tasks_lock:
                    tasks[task_id]['failed'] += 1
                mdb.log_entry(email, 'custom', subject, f'failed: {e}')
            time.sleep(0.1)

        # cleanup attachments after sending
        for p in attachments or []:
            try:
                os.remove(p)
            except Exception:
                pass
            try:
                parent = os.path.dirname(p)
                shutil.rmtree(parent, ignore_errors=True)
            except Exception:
                pass

        with tasks_lock:
            tasks[task_id]['status'] = 'done'
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        logging.exception('Custom send task failed (task_id=%s)', task_id)
        with tasks_lock:
            tasks[task_id]['status'] = 'error'
            tasks[task_id]['error'] = tb


@app.route('/greetings', methods=['GET'])
def greetings_form():
    return render_template('greetings.html', products=PRODUCTS)


@app.route('/greetings/send', methods=['POST'])
def greetings_send():
    # Bulk preview: accept an uploaded recipients file OR a single email, optional attachments, subject and body
    file = request.files.get('file')
    single_email = request.form.get('email')
    kind = request.form.get('type') or 'followup'
    subject = request.form.get('subject') or ''
    body = request.form.get('body') or ''
    body = sanitize_text_field(body)
    note = request.form.get('note') or ''
    product_key = request.form.get('product_key') or ''
    market_class = request.form.get('market_class') or ''  # e.g., k12, higher_ed, coaching

    # attachments uploaded by user (for preview -> will be saved and sent on confirm)
    attachments = request.files.getlist('attachments')

    # collect recipients (email + optional name)
    recipients = []

    tmpdir = None
    try:
        if file and file.filename:
            tmpdir = tempfile.mkdtemp()
            path = os.path.join(tmpdir, secure_filename(file.filename))
            file.save(path)
            try:
                if path.lower().endswith('.csv'):
                    df = pd.read_csv(path)
                else:
                    df = pd.read_excel(path, engine='openpyxl')
            except Exception as e:
                flash('Failed to read recipients file: ' + str(e), 'danger')
                return redirect(request.url)

            emails = extract_emails_from_dataframe(df)
            # Try to find a name column
            name_col = None
            for c in df.columns:
                if 'name' in str(c).lower() or 'first' in str(c).lower() or 'full' in str(c).lower():
                    name_col = c
                    break
            if name_col:
                for e in emails:
                    # best-effort find row where email matches
                    matches = df[df.apply(lambda r: any(str(v).lower().find(str(e).lower()) != -1 for v in r.values), axis=1)]
                    name_val = ''
                    if not matches.empty and name_col in matches.columns:
                        try:
                            name_val = str(matches.iloc[0][name_col])
                        except Exception:
                            name_val = ''
                    recipients.append({'email': e, 'name': name_val})
            else:
                recipients = [{'email': e, 'name': ''} for e in emails]
        elif single_email:
            name = request.form.get('name') or ''
            recipients = [{'email': single_email.strip(), 'name': name}]
        else:
            flash('Please provide recipients via upload or single email', 'warning')
            return redirect(request.url)

        if not subject:
            subject = f"{kind.title()} from EduAI"

        # Save attachments temporarily for the preview->confirm flow
        attach_paths = []
        if attachments:
            attach_tmp = tempfile.mkdtemp()
            for f in attachments:
                if getattr(f, 'filename', None):
                    path = os.path.join(attach_tmp, secure_filename(f.filename))
                    f.save(path)
                    attach_paths.append(path)

        # Render a preview using the body provided.
        # If the user left the body empty, generate a full AI-written greeting (regardless of the "use AI" checkbox)
        # Otherwise, when a body is provided, optionally rewrite it using OpenAI, or fall back to the local stylizer.
        if not body:
            # Generate a full greeting using the greetings helper (AI) but request fragments
            product_name = None
            product_pains = None
            if product_key:
                p = PRODUCTS.get(product_key)
                if p:
                    product_name = p.get('name')
                    product_pains = p.get('pains')
            try:
                subj_ai, greeting_frag, main_frag, closing_frag = mgreet.generate_greeting(kind, recipients[0].get('name', ''), '', product_name=product_name, product_pains=product_pains, return_fragments=True)
                # build a single fragment to store/send
                body_fragment = (greeting_frag or '') + (main_frag or '') + (closing_frag or '')
                # choose template by market_class when available
                tmpl = f"email_template_{market_class}.html" if market_class else 'email_template_greeting.html'
                try:
                    body_html = render_template(tmpl, recipient_name='', ai_hook=subj_ai or '', ai_greeting=greeting_frag, ai_main_body=main_frag, ai_closing=closing_frag)
                except Exception:
                    # fallback if market-specific template not present
                    body_html = render_template('email_template_greeting.html', recipient_name='', ai_hook=subj_ai or '', ai_greeting=greeting_frag, ai_main_body=main_frag, ai_closing=closing_frag)
            except Exception:
                logging.exception('Greeting generation failed; using local fallback')
                flash('AI greeting generation is unavailable — using a local pain-first fallback.', 'warning')
                # Fallback: craft a concise pain-first lead using known product pains if available
                if product_pains:
                    pains_snippet = ' '.join([f"{p}." for p in product_pains[:2]])
                    fallback_text = f"Struggling with {', '.join([p.split()[0] for p in product_pains[:2]])}? {pains_snippet} Try EduAIHub's tools to save time."
                else:
                    fallback_text = "Struggling with lesson planning and grading? Try EduAIHub's tools to save hours."
                body_fragment = stylize_marketing_body(fallback_text, product_name=product_name)
                # use market template if available
                tmpl = f"email_template_{market_class}.html" if market_class else 'email_template_greeting.html'
                try:
                    body_html = render_template(tmpl, recipient_name='', ai_hook='', ai_greeting='', ai_main_body=body_fragment, ai_closing=note)
                except Exception:
                    body_html = render_template('email_template_greeting.html', recipient_name='', ai_hook='', ai_greeting='', ai_main_body=body_fragment, ai_closing=note)

            if PREMAILER_AVAILABLE:
                # Keep style tags in preview so CSS pulse animation can render in the browser preview
                body_html = inline_css(body_html, keep_style_tags=True)
            else:
                flash('Premailer not installed — styles may not appear in some email clients. Install with `pip install premailer`.', 'warning')

        else:
            # If user asked to use AI to rewrite the body
            raw_use_ai = request.form.get('use_ai')
            use_ai = (raw_use_ai == '1') if raw_use_ai is not None else OPENAI_AVAILABLE
            if use_ai:
                try:
                    subj_ai, rewritten = mcustom.rewrite_body(body, recipient_name='', product_name=product_key and PRODUCTS.get(product_key, {}).get('name'), product_pains=PRODUCTS.get(product_key, {}).get('pains') if product_key else None)
                    # If subject not provided by user, use AI subject suggestion
                    if not subject and subj_ai:
                        subject = subj_ai
                    body_fragment = rewritten
                except Exception as e:
                    logging.exception('AI rewrite failed; falling back to local stylize')
                    flash('AI rewrite failed; using local stylize instead.', 'warning')
                    body_fragment = stylize_marketing_body(body, product_name=product_key and PRODUCTS.get(product_key, {}).get('name'))
            else:
                body_fragment = stylize_marketing_body(body, product_name=product_key and PRODUCTS.get(product_key, {}).get('name'))

            # Do NOT inject recipient name when user-provided body is used (avoids duplicate greetings)
            with app.app_context():
                tmpl = f"email_template_{market_class}.html" if market_class else 'email_template_greeting.html'
                try:
                    body_html = render_template(tmpl, recipient_name='', ai_hook='', ai_greeting='', ai_main_body=body_fragment, ai_closing=note)
                except Exception:
                    body_html = render_template('email_template_greeting.html', recipient_name='', ai_hook='', ai_greeting='', ai_main_body=body_fragment, ai_closing=note)
            if PREMAILER_AVAILABLE:
                # Keep style tags in preview so the pulse animation shows in the browser preview
                body_html = inline_css(body_html, keep_style_tags=True)
            else:
                flash('Premailer not installed — styles may not appear in some email clients. Install with `pip install premailer`.', 'warning')

        # Build recipient_data hidden payload (email||name per line)
        recipient_data = '\n'.join([f"{r['email']}||{r['name']}" for r in recipients])

        return render_template('preview.html', emails=[r['email'] for r in recipients], count=len(recipients), subject=subject, sender_name='EduAI', email_html=body_html, greeting_mode='greeting', greeting_kind=kind, custom_attachments='||'.join(attach_paths), recipient_data=recipient_data, email_fragment=body_fragment, structure_html='', structure_css='')
    finally:
        # Note: we keep attach_tmp and tempdir alive for the confirm step; they will be cleaned up after sending
        pass


@app.route('/greetings-start-send', methods=['POST'])
def greetings_start_send():
    # Accept recipient_data (email||name per line) and attachments list
    recipient_data_raw = request.form.get('recipient_data') or ''
    subject = request.form.get('subject') or 'Greeting from EduAI'
    kind = request.form.get('greeting_kind') or 'followup'
    sender_name = request.form.get('sender_name') or 'EduAI'
    attachments_raw = request.form.get('custom_attachments') or ''
    attachments = [p for p in attachments_raw.split('||') if p]
    # Prefer fragment (no header/footer) when available to render per-recipient; otherwise fall back to full html
    email_fragment = request.form.get('email_fragment') or ''
    email_html = request.form.get('email_html') or ''

    recipients = []
    for line in recipient_data_raw.splitlines():
        if not line.strip():
            continue
        parts = line.split('||')
        email = parts[0].strip()
        name = parts[1].strip() if len(parts) > 1 else ''
        recipients.append({'email': email, 'name': name})

    if not recipients:
        flash('No recipients to send to', 'warning')
        return redirect(url_for('greetings_form'))

    if request.form.get('dry_run') == '1':
        flash('Dry run: no emails were sent. Preview only.', 'info')
        results = {'sent': 0, 'failed': []}
        return render_template('result.html', results=results)

    # Start background send
    task_id = str(uuid.uuid4())
    with tasks_lock:
        tasks[task_id] = {'status': 'pending', 'total': len(recipients), 'sent': 0, 'failed': 0}

    # Pass the fragment into the background task so it's available during rendering
    thread = threading.Thread(target=send_greetings_task, args=(task_id, recipients, subject, sender_name, attachments, email_html, email_fragment, kind), daemon=True)
    thread.start()
    return redirect(url_for('progress', task_id=task_id))


def send_greetings_task(task_id, recipients, subject, sender_display=None, attachments=None, body_html=None, email_fragment=None, kind='followup'):
    try:
        with tasks_lock:
            tasks[task_id]['status'] = 'running'
            tasks[task_id]['total'] = len(recipients)
            tasks[task_id]['sent'] = 0
            tasks[task_id]['failed'] = 0

        for r in recipients:
            email = r.get('email')
            name = r.get('name')
            with app.app_context():
                # If we have a fragment (preferred), render the template per-recipient so name can be inserted
                fragment_to_use = email_fragment if email_fragment else body_html

                # Sanitize fragment_to_use to avoid nested headers/footers if it accidentally contains a full page
                if fragment_to_use and ("<html" in fragment_to_use.lower() or "<header" in fragment_to_use.lower() or "<h1" in fragment_to_use.lower() or "eduaihub" in fragment_to_use.lower()):
                    # remove common full-page blocks and company header/footer
                    fragment_to_use = re.sub(r'(?is)<\s*head[^>]*>.*?<\s*/\s*head\s*>', '', fragment_to_use)
                    fragment_to_use = re.sub(r'(?is)<\s*html[^>]*>', '', fragment_to_use)
                    fragment_to_use = re.sub(r'(?is)<\s*/\s*html\s*>', '', fragment_to_use)
                    fragment_to_use = re.sub(r'(?is)<\s*body[^>]*>', '', fragment_to_use)
                    fragment_to_use = re.sub(r'(?is)<\s*/\s*body\s*>', '', fragment_to_use)
                    fragment_to_use = re.sub(r'(?is)<\s*header[^>]*>.*?<\s*/\s*header\s*>', '', fragment_to_use)
                    fragment_to_use = re.sub(r'(?is)<\s*h1[^>]*>.*?EduAIHub.*?<\s*/\s*h1\s*>', '', fragment_to_use)
                    fragment_to_use = re.sub(r'(?is)<\s*footer[^>]*>.*?<\s*/\s*footer\s*>', '', fragment_to_use)
                    fragment_to_use = re.sub(r'(?is)<a[^>]*href=["\"][^"\"]*eduaihub[^"\"]["\"][^>]*>.*?<\s*/\s*a\s*>', '', fragment_to_use)
                    fragment_to_use = re.sub(r'\n{3,}', '\n\n', fragment_to_use).strip()

                rendered = render_template('email_template_greeting.html', recipient_name=name, ai_hook='', ai_greeting='', ai_main_body=fragment_to_use, ai_closing='')
                if PREMAILER_AVAILABLE:
                    # Keep style tags in the outgoing email so clients that respect <style> can show the pulse animation.
                    rendered = inline_css(rendered, keep_style_tags=True)

            ok, err = mutils.send_email_with_attachments(email, subject, rendered, attachments=attachments or None, sender_name=sender_display)
            if ok:
                with tasks_lock:
                    tasks[task_id]['sent'] += 1
                mdb.log_entry(email, f'greeting-{kind}', subject, 'sent')
            else:
                with tasks_lock:
                    tasks[task_id]['failed'] += 1
                mdb.log_entry(email, f'greeting-{kind}', subject, f'failed: {err}')
            time.sleep(0.1)

        # cleanup attachments after sending
        for p in attachments or []:
            try:
                os.remove(p)
            except Exception:
                pass
            try:
                parent = os.path.dirname(p)
                shutil.rmtree(parent, ignore_errors=True)
            except Exception:
                pass

        with tasks_lock:
            tasks[task_id]['status'] = 'done'
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        logging.exception('Greeting send task failed (task_id=%s)', task_id)
        with tasks_lock:
            tasks[task_id]['status'] = 'error'
            tasks[task_id]['error'] = tb


@app.route('/logs')
def logs_page():
    rows = mdb.get_logs(500)
    return render_template('logs.html', rows=rows)


@app.route('/bulk', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('file')
        subject = request.form.get('subject') or 'Explore EduAIHub — AI tools for classrooms'
        sender_name = request.form.get('sender_name')
        if not file or file.filename == '':
            flash('No file selected', 'danger')
            return redirect(request.url)
        if not allowed_file(file.filename):
            flash('Unsupported file type', 'danger')
            return redirect(request.url)

        filename = secure_filename(file.filename)
        tmpdir = tempfile.mkdtemp()
        path = os.path.join(tmpdir, filename)
        file.save(path)

        try:
            if filename.lower().endswith('.csv'):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path, engine='openpyxl')
        except Exception as e:
            flash('Failed to read uploaded file: ' + str(e), 'danger')
            return redirect(request.url)

        emails = extract_emails_from_dataframe(df)
        if not emails:
            flash('No email addresses found in the file', 'warning')
            return redirect(request.url)

        intro = sanitize_text_field(request.form.get('intro') or '')
        # choose a banner image (use first product image if available)
        banner_url = None
        for v in PRODUCTS.values():
            if v.get('image_url'):
                banner_url = v.get('image_url')
                break
        email_html = render_template('email_template_bulk.html', products=PRODUCTS, intro=intro, banner_url=banner_url)
        # try to inline CSS so preview matches how email clients render
        if PREMAILER_AVAILABLE:
            # Keep style tags in preview so pulse animation shows in the browser preview
            email_html = inline_css(email_html, keep_style_tags=True)
        else:
            flash('Premailer not installed — styles may not appear in some email clients. Install with `pip install premailer`.', 'warning')
        return render_template('preview.html', emails=emails, count=len(emails), subject=subject, sender_name=sender_name or '', email_html=email_html, intro=intro, structure_html='', structure_css='')    
    return render_template('index.html')


# Product-wise sender
PRODUCTS = {
    'class_tom': {
        'name': 'Class Tom',
        'desc': 'Revolutionize Any Classroom with AI — turns ordinary classrooms into intelligent, interactive learning spaces without extra hardware.',
        'link': 'https://www.eduaihub.in/class-tom/',
        'image_url': 'https://www.eduaihub.in/wp-content/uploads/2025/03/class_tom.jpg',
        'pains': [
            'Lack of affordable smart classroom solutions',
            'High setup costs and hardware dependencies',
            'Limited interactivity with traditional teaching tools'
        ],
        'features': [
            'Works offline — no internet required',
            'No additional hardware needed',
            'Teacher-friendly integrations and analytics'
        ]
    },

    'vidya_hub': {
        'name': 'Vidya Hub',
        'desc': 'Teachers spend less time on paperwork and more time inspiring students — let AI handle grading, feedback, and planning.',
        'link': 'https://www.eduaihub.in/vidya-hub/',
        'image_url': 'https://www.eduaihub.in/wp-content/uploads/2025/03/vidya_hub.jpg',
        'pains': [
            'Overwhelming administrative workload for teachers',
            'Time-consuming grading and feedback processes',
            'Difficulty personalizing student feedback at scale'
        ],
        'features': [
            'Automated grading and feedback',
            'Lesson planning assistance',
            'Customizable teacher workflows'
        ]
    },
    'ai_viz_lab': {
        'name': 'AI Viz Lab',
        'desc': 'Learning reimagined with AI and creativity — hands-on visual labs for students to explore and create.',
        'link': 'https://www.eduaihub.in/vidya-hub-2/',
        'image_url': 'https://www.eduaihub.in/wp-content/uploads/2025/03/ai_viz_lab.jpg',
        'pains': [
            'Lack of engaging, hands-on AI learning activities for students',
            'Limited opportunities for experimentation in traditional curricula',
            'Insufficient tools for visual learning and creativity with AI'
        ],
        'features': [
            'Interactive visual AI experiments',
            'Project-based learning modules',
            'Easy integration with classroom workflows'
        ]
    }
}


@app.route('/product-sender', methods=['GET', 'POST'])
def product_sender():
    if request.method == 'POST':
        product_key = request.form.get('product_key')
        product = PRODUCTS.get(product_key)
        subject = request.form.get('subject') or f"Discover {product['name']} — from EduAI"
        sender_name = request.form.get('sender_name')
        custom_note = request.form.get('custom_note') or ''
        file = request.files.get('file')
        if not file or file.filename == '':
            flash('No file selected', 'danger')
            return redirect(request.url)
        if not allowed_file(file.filename):
            flash('Unsupported file type', 'danger')
            return redirect(request.url)

        filename = secure_filename(file.filename)
        tmpdir = tempfile.mkdtemp()
        path = os.path.join(tmpdir, filename)
        file.save(path)

        try:
            if filename.lower().endswith('.csv'):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path, engine='openpyxl')
        except Exception as e:
            flash('Failed to read uploaded file: ' + str(e), 'danger')
            return redirect(request.url)

        emails = extract_emails_from_dataframe(df)
        if not emails:
            flash('No email addresses found in the file', 'warning')
            return redirect(request.url)

        # sanitize note and build image URL and features
        custom_note = sanitize_text_field(custom_note)
        product_image = product.get('image_url')
        product_features = product.get('features', [])
        product_pains = product.get('pains', [])

        email_html = render_template('email_template_product.html', product_name=product['name'], product_desc=product['desc'], product_link=product['link'], product_image=product_image, product_features=product_features, product_pains=product_pains, custom_note=custom_note)
        if PREMAILER_AVAILABLE:
            # Keep style tags in preview so CSS pulse can render
            email_html = inline_css(email_html, keep_style_tags=True)
        else:
            flash('Premailer not installed — styles may not appear in some email clients. Install with `pip install premailer`.', 'warning')
        return render_template('preview.html', emails=emails, count=len(emails), subject=subject, sender_name=sender_name or '', email_html=email_html, product_key=product_key, custom_note=custom_note, product_name=product['name'], structure_html='', structure_css='')
    return render_template('product_send.html')


@app.route('/send', methods=['POST'])
def send():
    data = request.form
    emails_raw = data.get('emails') or ''
    subject = data.get('subject') or 'Explore EduAIHub — AI tools for classrooms'
    sender_name = data.get('sender_name')
    emails = [e.strip() for e in emails_raw.splitlines() if e.strip()]
    html_body = render_template('email_template_bulk.html')

    try:
        results = send_bulk_emails(emails, subject, html_body, sender_display=sender_name)
        return render_template('result.html', results=results)
    except Exception as e:
        flash('Error sending emails: ' + str(e), 'danger')
        return redirect(url_for('index'))


@app.route('/start-send', methods=['POST'])
def start_send():
    emails_raw = request.form.get('emails') or ''
    subject = request.form.get('subject') or 'Explore EduAIHub — AI tools for classrooms'
    sender_name = request.form.get('sender_name')
    product_key = request.form.get('product_key')
    intro = request.form.get('intro') or ''
    custom_note = request.form.get('custom_note') or ''
    emails = [e.strip() for e in emails_raw.splitlines() if e.strip()]
    if not emails:
        flash('No recipients provided', 'danger')
        return redirect(url_for('index'))

    # Render the same template as preview (bulk or product)
    if product_key:
        product = PRODUCTS.get(product_key)
        # render inside app_context; optionally use OpenAI to suggest a subject/hook when available
        ai_hook = ''
        if OPENAI_AVAILABLE:
            try:
                subj_ai, body_frag = mcustom.rewrite_body(product.get('desc', ''), recipient_name='', product_name=product.get('name'), product_pains=product.get('pains'))
                if subject == f"Discover {product['name']} — from EduAI":
                    subject = subj_ai or subject
                # extract a short hook from the AI fragment
                plain = re.sub(r'<[^>]+>', '', body_frag).strip()
                first_sent = re.split(r'(?<=[.!?])\s+', plain)[0] if plain else ''
                ai_hook = first_sent[:160]
            except Exception:
                ai_hook = ''
        # render inside app_context
        with app.app_context():
            html_body = render_template('email_template_product.html', product_name=product['name'], product_desc=product['desc'], product_link=product['link'], product_image=product.get('image_url'), product_features=product.get('features', []), product_pains=product.get('pains', []), custom_note=custom_note, ai_hook=ai_hook)
    else:
        with app.app_context():
            html_body = render_template('email_template_bulk.html', products=PRODUCTS, intro=intro)

    # inline CSS so sent message matches preview
    # Keep style tags to maximize chance that clients that respect <style> can render our pulse animation
    html_body = inline_css(html_body, keep_style_tags=True)

    # If dry_run flag provided, do not send — just show a dry-run result
    if request.form.get('dry_run') == '1':
        flash('Dry run: no emails were sent. Preview only.', 'info')
        results = {'sent': 0, 'failed': []}
        return render_template('result.html', results=results)

    task_id = str(uuid.uuid4())
    with tasks_lock:
        tasks[task_id] = {'status': 'pending', 'total': len(emails), 'sent': 0, 'failed': 0}

    thread = threading.Thread(target=send_task, args=(task_id, emails, subject, html_body, sender_name), daemon=True)
    thread.start()
    return redirect(url_for('progress', task_id=task_id))

@app.route('/product-start-send', methods=['POST'])
def product_start_send():
    emails_raw = request.form.get('emails') or ''
    product_key = request.form.get('product_key')
    product = PRODUCTS.get(product_key)
    if not product:
        flash('Invalid product selected', 'danger')
        return redirect(url_for('product_sender'))
    subject = request.form.get('subject') or f"Discover {product['name']} — from EduAI"
    sender_name = request.form.get('sender_name')
    custom_note = request.form.get('custom_note') or ''
    emails = [e.strip() for e in emails_raw.splitlines() if e.strip()]
    if not emails:
        flash('No recipients provided', 'danger')
        return redirect(url_for('product_sender'))

    # Support dry-run from preview
    if request.form.get('dry_run') == '1':
        flash('Dry run: no emails were sent. Preview only.', 'info')
        results = {'sent': 0, 'failed': []}
        return render_template('result.html', results=results)

    task_id = str(uuid.uuid4())
    with tasks_lock:
        tasks[task_id] = {'status': 'pending', 'total': len(emails), 'sent': 0, 'failed': 0}

    thread = threading.Thread(target=send_product_task, args=(task_id, emails, subject, sender_name, product_key, custom_note), daemon=True)
    thread.start()
    return redirect(url_for('progress', task_id=task_id))


def send_product_task(task_id, recipients, subject, sender_display=None, product_key=None, custom_note=''):
    """Background task to send product emails and log each send. Embeds product image inline when available."""
    try:
        with tasks_lock:
            tasks[task_id]['status'] = 'running'
            tasks[task_id]['total'] = len(recipients)
            tasks[task_id]['sent'] = 0
            tasks[task_id]['failed'] = 0

        product = PRODUCTS.get(product_key) if product_key else None

        for r in recipients:
            # Use remote image URL if available
            product_image_ref = product.get('image_url') if product and product.get('image_url') else None
            # render template inside application context because this runs in a background thread
            with app.app_context():
                html_body = render_template(
                    'email_template_product.html',
                    product_name=product['name'],
                    product_desc=product['desc'],
                    product_link=product['link'],
                    product_image=product_image_ref,
                    product_features=product.get('features', []),
                    product_pains=product.get('pains', []),
                    custom_note=custom_note,
                )
                # Keep style tags in outgoing product emails to improve client support for the pulse animation
                html_body = inline_css(html_body, keep_style_tags=True)

            ok, err = mutils.send_email_with_attachments(r, subject, html_body, attachments=None, inline_images=None, sender_name=sender_display)
            if ok:
                with tasks_lock:
                    tasks[task_id]['sent'] += 1
                mdb.log_entry(r, f'product-{product_key}', subject, 'sent')
            else:
                with tasks_lock:
                    tasks[task_id]['failed'] += 1
                mdb.log_entry(r, f'product-{product_key}', subject, f'failed: {err}')
            time.sleep(0.1)

        with tasks_lock:
            tasks[task_id]['status'] = 'done'
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        logging.exception('Product send task failed (task_id=%s)', task_id)
        with tasks_lock:
            tasks[task_id]['status'] = 'error'
            tasks[task_id]['error'] = tb

@app.route('/progress/<task_id>')
def progress(task_id):
    with tasks_lock:
        task = tasks.get(task_id)
    if not task:
        flash('Task not found', 'danger')
        return redirect(url_for('index'))
    return render_template('progress.html', task_id=task_id, total=task.get('total', 0))


@app.route('/status/<task_id>')
def status(task_id):
    with tasks_lock:
        task = tasks.get(task_id)
        if not task:
            return jsonify({'ok': False, 'error': 'task not found'}), 404
        sent = task.get('sent', 0)
        failed = task.get('failed', 0)
        total = task.get('total', 0)
        status_val = task.get('status', 'pending')
    remaining = max(0, total - (sent + failed))
    response = {'ok': True, 'status': status_val, 'sent': sent, 'failed': failed, 'total': total, 'remaining': remaining}
    # include error trace if present
    if 'error' in task:
        response['error'] = task.get('error')
    return jsonify(response)


@app.route('/final/<task_id>')
def final(task_id):
    with tasks_lock:
        task = tasks.get(task_id)
    if not task:
        flash('Task not found', 'danger')
        return redirect(url_for('index'))
    # Show success counts but do not list failed addresses per user request
    return render_template('final.html', total=task.get('total', 0), sent=task.get('sent', 0), failed=task.get('failed', 0))


@app.route('/task/<task_id>')
def task_debug(task_id):
    """Render a debug page for a background task (status, error trace, and recent logs)."""
    with tasks_lock:
        task = tasks.get(task_id)
    if not task:
        flash('Task not found', 'danger')
        return redirect(url_for('index'))

    # Fetch recent logs to help debugging
    rows = mdb.get_logs(200)
    return render_template('task_debug.html', task_id=task_id, task=task, rows=rows)


if __name__ == '__main__':
    # initialize DB
    mdb.init_db()
    app.run(debug=True)

