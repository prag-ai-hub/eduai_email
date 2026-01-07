import os
import re
from flask import render_template

# Import OpenAI lazily and safely so the app doesn't fail to import when openai is
# not installed or the API key is missing. If OpenAI is not available, the functions
# below will raise a RuntimeError which will let the caller fall back to a local
# stylizer.
try:
    import openai
    openai.api_key = os.environ.get('OPENAI_API_KEY')
except Exception:
    openai = None


def generate_custom_email(recipient_name, description, product_name=None, product_pains=None, company_name='EduAIHub'):
    """Generate a subject, hook and HTML body using OpenAI based on description and recipient name.
    Returns (subject, rendered_html)
    """
    pains_text = ''
    if product_name and product_pains:
        pains_text = f"Product: {product_name}. Known pains: {', '.join(product_pains)}."

    prompt = (
        f"You are a professional marketing copywriter for {company_name}.\n"
        f"Use a pain-first marketing approach: start by describing the key problems or pain points the recipient faces (use the product pains if provided), then present the product benefits as direct solutions to those problems.\n"
        f"Produce a short subject line (one sentence), a one-line compelling hook (ideally pain-oriented), and an HTML body section (no full html page) suitable to embed into an existing email template.\n"
        f"Recipient name: {recipient_name if recipient_name else 'Valued Educator'}.\n"
        f"{pains_text}\n"
        f"Context / description: {description}\n"
        "Respond with the format:\nSubject: <subject>\nHook: <short hook>\nBody:\n<html>...</html>\n"
    )

    if not openai or not getattr(openai, 'api_key', None):
        raise RuntimeError('OpenAI is not available; set OPENAI_API_KEY and install the openai package to use AI features')
    # Use a wrapper that supports both older and newer openai client interfaces
    def _chat_complete(messages, model='gpt-3.5-turbo', **kwargs):
        if not openai or not getattr(openai, 'api_key', None):
            raise RuntimeError('OpenAI is not available; set OPENAI_API_KEY and install the openai package to use AI features')
        try:
            return openai.ChatCompletion.create(model=model, messages=messages, **kwargs)
        except Exception:
            # Newer client provides the chat completions under openai.chat.completions
            try:
                return openai.chat.completions.create(model=model, messages=messages, **kwargs)
            except Exception as e:
                raise

    resp = _chat_complete([{'role': 'system', 'content': 'You write concise marketing emails.'}, {'role': 'user', 'content': prompt}], max_tokens=500, temperature=0.6)

    text = resp.choices[0].message.content.strip()

    subject = f'Update from {company_name}'
    hook = ''
    body_html = text
    # improved parsing: remove Hook: and Body: labels from body
    lines = text.splitlines()
    if lines and lines[0].lower().startswith('subject:'):
        subject = lines[0].split(':', 1)[1].strip()
        lines = lines[1:]

    cleaned_lines = []
    for line in lines:
        if not line.strip():
            cleaned_lines.append('')
            continue
        low = line.strip().lower()
        if low.startswith('hook:'):
            try:
                hook = line.split(':', 1)[1].strip()
            except Exception:
                hook = ''
            continue
        if low.startswith('body:'):
            # drop the 'Body:' label but keep any trailing content on that line
            rest = line.split(':', 1)[1].strip()
            if rest:
                cleaned_lines.append(rest)
            continue
        cleaned_lines.append(line)

    body_html = '\n'.join([l for l in cleaned_lines if l is not None])

    # Clean AI output: remove repeated headers, greetings and footers so the fragment can be safely
    # embedded into the main template without producing duplicates.
    # Remove leading greetings like "Hi Abhishek," or "Dear Abhishek,"
    body_html = re.sub(r'(?i)^\s*(hi|hello|dear)\s+[^\n,]{1,80},?\s*\n', '', body_html)
    # Remove common signature/footer lines and repeated company headers
    body_html = re.sub(r'(?is)\n?\s*(warm regards|regards|thanks|thank you)[^\n]*', '', body_html)
    body_html = re.sub(r'(?is)\n?\s*(visit\s+eduaihub[^\n]*|unsubscribe[^\n]*|visit[^\n]*eduaihub[^\n]*)', '', body_html)
    body_html = re.sub(r"(?im)^(\s*EduAIHub\s*\n\s*Practical AI Tools for Education\s*\n)+", '', body_html)
    body_html = re.sub(r"(?im)^(\s*EduAIHub\s*\n)+", '', body_html)
    # Strip common html header/footer tags if present
    body_html = re.sub(r'(?is)<\s*header[^>]*>.*?<\s*/\s*header\s*>', '', body_html)
    body_html = re.sub(r'(?is)<\s*footer[^>]*>.*?<\s*/\s*footer\s*>', '', body_html)
    # Remove any remaining lines that look like the company footer to be safe
    body_html = re.sub(r'(?im)^.*EduAI\s*Hub.*$', '', body_html, flags=re.MULTILINE)
    # Trim excess whitespace and blank lines
    body_html = re.sub(r'\n{3,}', '\n\n', body_html).strip()

    # Return subject, cleaned HTML fragment and hook (caller will render it into the full template)
    return subject, body_html, hook


def rewrite_body(raw_body: str, recipient_name: str = '', product_name: str | None = None, product_pains=None, company_name='EduAIHub', structure_only: bool = False) -> tuple:
    """Use OpenAI to rewrite a provided body into a pain-first marketing HTML fragment and suggest a subject.
    Returns (subject, html_fragment)"""
    pains_text = ''
    if product_name and product_pains:
        pains_text = f"Product: {product_name}. Known pains: {', '.join(product_pains)}."

    def _local_structure_format(text: str) -> str:
        if not text:
            return ''
        # Replace greeting names with placeholder
        text = re.sub(r'(?im)^(\s*(hi|hello|dear)\s+)([^\n,]{1,80})(,?)', lambda m: f"{m.group(1)}[[RECIPIENT_NAME]]{m.group(4) or ','}", text)
        # Normalize blank lines
        text = re.sub(r'\n{2,}', '\n\n', text).strip()
        paras = [p.strip() for p in text.split('\n\n') if p.strip()]
        p_style = "margin:0 0 12px 0; font-size:15px; color:#234b38; line-height:1.6; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif"
        if paras:
            first = paras[0]
            parts = re.split(r'(?<=[.!?])\s+', first, maxsplit=1)
            if parts:
                lead = f"<strong>{parts[0]}</strong>" + ((' ' + parts[1]) if len(parts) > 1 else '')
                paras[0] = lead
        return ''.join([f"<p style=\"{p_style}\">{p}</p>" for p in paras])

    # If the caller asked only for structure (no content changes), do that and return early
    if structure_only:
        # Remove any pasted instruction/code blocks (triple-backtick fenced blocks) and simple echo of instruction lines
        raw_body_sanitized = re.sub(r'```.*?```', '', raw_body, flags=re.S).strip()
        raw_body_sanitized = re.sub(r'(?im)^.*you are a professional.*$', '', raw_body_sanitized).strip()
        # If the user pasted an instruction block (contains directives like 'Do NOT change' or 'Respond with only'),
        # strip leading instruction-like paragraphs (bullets, directives) until we find a paragraph that looks like the actual message content.
        if re.search(r'(?is)do not change|respond with only|you are a professional', raw_body_sanitized):
            paras = re.split(r'\n\s*\n', raw_body_sanitized)
            keep_index = 0
            for i, p in enumerate(paras):
                lp = p.strip().lower()
                # treat as instruction if it contains common directive phrases or is a bullet list
                if lp.startswith('-') or re.search(r'(?i)do not change|respond with only|preserve all links|replace any recipient|format the provided message|wrap text|bold only|do not paraphrase', lp):
                    continue
                # otherwise we consider it message content
                keep_index = i
                break
            if keep_index < len(paras):
                raw_body_sanitized = '\n\n'.join(paras[keep_index:]).strip()

        if openai and getattr(openai, 'api_key', None):
            prompt_structure = (
                f"You are a professional email formatter for {company_name}.\n"
                "Do NOT change the user's words or meaning. Only format the provided message into clean, accessible HTML suitable for email:\n"
                " - Wrap text into short paragraphs and add minimal inline styling.\n"
                " - Bold only the first sentence (problem lead).\n"
                " - Preserve all links and images.\n"
                " - Do NOT paraphrase, add, or remove sentences.\n"
                " - Replace any recipient names in greetings (e.g., 'Hi Support,', 'Dear Team,') with the exact token [[RECIPIENT_NAME]] so the sending code can personalize per recipient.\n"
                "Respond with only an HTML fragment (no full <html> document).\n"
                "Original message:\n" + raw_body_sanitized + "\n"
            )
            try:
                def _chat_complete(messages, model='gpt-3.5-turbo', **kwargs):
                    try:
                        return openai.ChatCompletion.create(model=model, messages=messages, **kwargs)
                    except Exception:
                        return openai.chat.completions.create(model=model, messages=messages, **kwargs)

                resp = _chat_complete([{'role': 'system', 'content': 'You only format text into HTML, do not change wording.'}, {'role': 'user', 'content': prompt_structure}], max_tokens=400, temperature=0.0)
                text = resp.choices[0].message.content.strip()

                # If model accidentally echoed the instruction prompt or returned the prompt itself, fall back to local formatter
                if 'you are a professional email formatter' in text.lower() or 'do not change the user' in text.lower() or text.strip().startswith('```'):
                    return '', _local_structure_format(raw_body_sanitized)

                lines = text.splitlines()
                if lines and lines[0].lower().startswith('subject:'):
                    subject = lines[0].split(':', 1)[1].strip()
                    body_text = '\n'.join(lines[1:]).strip()
                else:
                    subject = ''
                    body_text = text
                # Ensure greeting placeholder exists
                body_text = re.sub(r'(?im)^(\s*(hi|hello|dear)\s+)([^\n,]{1,80})(,?)', lambda m: f"{m.group(1)}[[RECIPIENT_NAME]]{m.group(4) or ','}", body_text)
                # If returned plaintext, wrap paragraphs locally
                if not re.search(r'(?i)<p\b', body_text):
                    body_text = _local_structure_format(body_text)
                return subject, body_text
            except Exception:
                return '', _local_structure_format(raw_body_sanitized)
        else:
            return '', _local_structure_format(raw_body_sanitized)

    prompt = (
        f"You are a professional marketing copywriter for {company_name}.\n"
        "Rewrite the following message into a concise, pain-first marketing email body suitable to embed into an existing email template. "
        "Start with a short, bold problem statement (one sentence), then 1-2 short paragraphs describing benefits and a single clear call to action. Keep paragraphs short, suitable for email. Preserve any links provided. Output only a subject line then the HTML fragment.\n"
        f"Recipient name: {recipient_name if recipient_name else 'Valued Educator'}.\n"
        f"{pains_text}\n"
        "Original content:\n" + raw_body + "\n"
        "Respond with the format:\nSubject: <subject>\nBody:\n<html>...</html>\n"
    )

    if not openai or not getattr(openai, 'api_key', None):
        raise RuntimeError('OpenAI is not available; set OPENAI_API_KEY and install the openai package to use AI features')
    # Use the same compatibility wrapper for chat completions
    def _chat_complete(messages, model='gpt-3.5-turbo', **kwargs):
        if not openai or not getattr(openai, 'api_key', None):
            raise RuntimeError('OpenAI is not available; set OPENAI_API_KEY and install the openai package to use AI features')
        try:
            return openai.ChatCompletion.create(model=model, messages=messages, **kwargs)
        except Exception:
            try:
                return openai.chat.completions.create(model=model, messages=messages, **kwargs)
            except Exception as e:
                raise

    resp = _chat_complete([{'role': 'system', 'content': 'You write short, effective marketing email copy.'}, {'role': 'user', 'content': prompt}], max_tokens=400, temperature=0.6)

    text = resp.choices[0].message.content.strip()

    subject = ''
    body_html = text
    lines = text.splitlines()
    if lines and lines[0].lower().startswith('subject:'):
        subject = lines[0].split(':', 1)[1].strip()
        lines = lines[1:]

    # Normalize lines and handle an explicit 'Body:' label (keep any trailing content on the same line)
    normalized = []
    # skip leading blank lines
    i = 0
    while i < len(lines) and not lines[i].strip():
        i += 1
    for line in lines[i:]:
        low = line.strip().lower()
        if low.startswith('body:'):
            rest = line.split(':', 1)[1].strip()
            if rest:
                normalized.append(rest)
            continue
        normalized.append(line)

    body_html = '\n'.join([l for l in normalized if l is not None]).strip()

    # Remove trailing signatures or repeated footers if model included them
    body_html = re.sub(r'(?is)\n?\s*(warm regards|regards|thanks|thank you)[^\n]*', '', body_html)
    body_html = re.sub(r'(?is)\n?\s*(visit\s+eduaihub[^\n]*|unsubscribe[^\n]*|visit[^\n]*eduaihub[^\n]*)', '', body_html)
    body_html = re.sub(r'(?is)\n?\s*\"[^\n]{0,200}\"\s*—\s*[^\n]{0,100}', '', body_html)

    # Remove any leading header lines the model may have included (text or simple lines)
    body_html = re.sub(r"(?im)^(\s*EduAIHub\s*\n\s*Practical AI Tools for Education\s*\n)+", '', body_html)
    body_html = re.sub(r"(?im)^(\s*EduAIHub\s*\n)+", '', body_html)

    # If model returned a full HTML document, strip outer tags and common header/footer blocks
    body_html = re.sub(r'(?is)<\s*head[^>]*>.*?<\s*/\s*head\s*>', '', body_html)
    body_html = re.sub(r'(?is)<\s*html[^>]*>', '', body_html)
    body_html = re.sub(r'(?is)<\s*/\s*html\s*>', '', body_html)
    body_html = re.sub(r'(?is)<\s*body[^>]*>', '', body_html)
    body_html = re.sub(r'(?is)<\s*/\s*body\s*>', '', body_html)
    body_html = re.sub(r'(?is)<\s*header[^>]*>.*?<\s*/\s*header\s*>', '', body_html)
    body_html = re.sub(r'(?is)<\s*h1[^>]*>.*?EduAIHub.*?<\s*/\s*h1\s*>', '', body_html)
    body_html = re.sub(r'(?is)<\s*footer[^>]*>.*?<\s*/\s*footer\s*>', '', body_html)
    body_html = re.sub(r'(?is)<a[^>]*href=["\"][^"\"]*eduaihub[^"\"]["\"][^>]*>.*?<\s*/\s*a\s*>', '', body_html)

    # Clean up extra blank lines and whitespace
    body_html = re.sub(r'\n{3,}', '\n\n', body_html).strip()

    # Add paragraph inline styles for consistency across clients
    p_style = "margin:0 0 12px 0; font-size:15px; color:#234b38; line-height:1.6; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif"
    # Add style attribute to <p> tags that don't already have one
    body_html = re.sub(r'(?i)<p\b(?![^>]*\bstyle=)([^>]*)>', lambda m: f"<p{m.group(1)} style=\"{p_style}\">", body_html)

    # If AI returned plaintext without <p> tags, wrap paragraphs
    if not re.search(r'(?i)<p\b', body_html):
        paras = [l.strip() for l in body_html.split('\n') if l.strip()]
        body_html = ''.join([f"<p style=\"{p_style}\">{p}</p>" for p in paras])

    # Add a subtle animated banner and a clear CTA to the fragment so AI output becomes more engaging.
    # Use ANIMATED_GIF_URL and CTA_PULSE_URL only when explicitly set in env. No hardcoded GIF fallbacks.
    default_gif = os.environ.get('ANIMATED_GIF_URL')
    # Add a style block defining a small pulse animation for clients that honor style tags. Keep pulse_style always available.
    pulse_style = (
        "<style>"
        ".pulse{display:inline-block;vertical-align:middle;margin-left:8px;border-radius:4px;}"
        "@keyframes pulse{0%{transform:scale(1);opacity:1}50%{transform:scale(1.12);opacity:0.9}100%{transform:scale(1);opacity:1}}"
        ".pulse-anim{animation:pulse 1.6s infinite ease-in-out;}"
        "</style>"
    )
    media_html = ''
    if default_gif:
        media_html = f"<div style=\"margin-bottom:12px;text-align:center;\"><img src=\"{default_gif}\" alt=\"\" width=320 style=\"display:block;border-radius:10px;max-width:100%;height:auto;\" /></div>"

    cta_link = os.environ.get('DEFAULT_CTA_LINK', 'https://www.eduaihub.in/demo')
    cta_icon = os.environ.get('CTA_PULSE_URL')
    # Use the pulse-anim CSS class only when CTA icon is explicitly provided; otherwise no icon.
    icon_html = f"<img src=\"{cta_icon}\" alt=\"\" width=18 class=\"pulse pulse-anim\" style=\"display:inline-block;vertical-align:middle;\" />" if cta_icon else ''
    cta_html = (
        f"<div style=\"margin-top:12px;text-align:center;\">"
        f"<a href=\"{cta_link}\" style=\"display:inline-block;padding:10px 16px;background:linear-gradient(90deg,#2fc071,#1aa35a);color:#fff;text-decoration:none;border-radius:8px;font-weight:700;box-shadow:0 6px 18px rgba(46,139,87,0.18);\">"
        f"Request a short demo » {icon_html}</a></div>"
    )

    # Prepend pulse style first (so it is always available) then media and append CTA to make the fragment more animated and action-oriented
    body_html = pulse_style + (media_html if media_html else '') + body_html + cta_html

    return subject, body_html
