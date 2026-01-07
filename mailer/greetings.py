import os
import re
from flask import render_template

# Import OpenAI safely so missing package or API key doesn't break app imports
try:
    import openai
    openai.api_key = os.environ.get('OPENAI_API_KEY')
except Exception:
    openai = None


def generate_greeting(kind, recipient_name, description, product_name=None, product_pains=None, return_fragments=False):
    """kind: invitation, thankyou, onboarding, followup
    Returns:
      - By default: (subject, rendered_html) where rendered_html is the full email template (safe to preview/send as-is)
      - If `return_fragments=True`: returns (subject, greeting_html, main_body_html, closing_html) suitable to be embedded into the email template without duplicating headers/footers
    """
    system = 'You are a professional email writer for the education sector. Write clear, warm, and engaging emails with proper structure. Use a pain-first approach where relevant.'

    pains_ctx = ''
    if product_name and product_pains:
        pains_ctx = f'Product: {product_name}. Known pains: {", ".join(product_pains)}. Use these pains in the introduction where relevant.'

    if kind == 'invitation':
        prompt = (
            f"Write a professional invitation email. Recipient: {recipient_name if recipient_name else 'Educator'}. "
            f"Meeting details: {description}\n"
            f"{pains_ctx}\n"
            f"Format your response as:\n"
            f"Subject: <subject>\n"
            f"Hook: <compelling one-liner>\n"
            f"Greeting: <1-2 sentences greeting paragraph>\n"
            f"Main Body: <2-3 paragraphs with meeting details, purpose, and next steps>\n"
            f"Closing: <1 sentence closing paragraph with call to action>"
        )
    elif kind == 'thankyou':
        prompt = (
            f"Write a warm thank-you email. Recipient: {recipient_name if recipient_name else 'Valued Colleague'}. "
            f"Context: {description}\n"
            f"{pains_ctx}\n"
            f"Format your response as:\n"
            f"Subject: <subject>\n"
            f"Hook: <heartfelt one-liner>\n"
            f"Greeting: <1-2 sentences greeting acknowledging their involvement>\n"
            f"Main Body: <2-3 paragraphs expressing gratitude, specific appreciation, and impact>\n"
            f"Closing: <1 sentence closing with forward-looking sentiment>"
        )
    elif kind == 'onboarding':
        prompt = (
            f"Write a warm onboarding welcome email. Recipient: {recipient_name if recipient_name else 'New Member'}. "
            f"Context: {description}\n"
            f"{pains_ctx}\n"
            f"Format your response as:\n"
            f"Subject: <subject>\n"
            f"Hook: <welcoming one-liner>\n"
            f"Greeting: <1-2 sentences welcoming them>\n"
            f"Main Body: <2-3 paragraphs with welcome message, what to expect, next steps, and support offer>\n"
            f"Closing: <1 sentence closing with encouragement>"
        )
    else:  # followup
        prompt = (
            f"Write a professional follow-up email. Recipient: {recipient_name if recipient_name else 'Valued Contact'}. "
            f"Context: {description}\n"
            f"{pains_ctx}\n"
            f"Format your response as:\n"
            f"Subject: <subject>\n"
            f"Hook: <engaging one-liner>\n"
            f"Greeting: <1-2 sentences greeting>\n"
            f"Main Body: <2-3 paragraphs with follow-up message, key points, and proposed next steps>\n"
            f"Closing: <1 sentence closing with call to action>"
        )

    if not openai or not getattr(openai, 'api_key', None):
        raise RuntimeError('OpenAI is not available; set OPENAI_API_KEY and install the openai package to use AI features')
    # Compatibility wrapper for older/newer OpenAI python clients
    def _chat_complete(messages, model='gpt-3.5-turbo', **kwargs):
        if not openai or not getattr(openai, 'api_key', None):
            raise RuntimeError('OpenAI is not available; set OPENAI_API_KEY and install the openai package to use AI features')
        try:
            return openai.ChatCompletion.create(model=model, messages=messages, **kwargs)
        except Exception:
            try:
                return openai.chat.completions.create(model=model, messages=messages, **kwargs)
            except Exception:
                raise

    response = _chat_complete([{'role': 'system', 'content': system}, {'role': 'user', 'content': prompt}], max_tokens=600, temperature=0.65)

    text = response.choices[0].message.content.strip()

    subject = f'{kind.title()} from EduAI'
    hook = ''
    greeting_section = ''
    main_body_section = ''
    closing_section = ''
    
    # Parse sections from response
    lines = text.splitlines()
    current_section = None
    section_content = []
    
    for line in lines:
        lower_line = line.lower().strip()
        if lower_line.startswith('subject:'):
            subject = line.split(':', 1)[1].strip()
        elif lower_line.startswith('hook:'):
            hook = line.split(':', 1)[1].strip()
        elif lower_line.startswith('greeting:'):
            if section_content and current_section:
                if current_section == 'greeting':
                    greeting_section = '\n'.join(section_content)
                elif current_section == 'main body':
                    main_body_section = '\n'.join(section_content)
                elif current_section == 'closing':
                    closing_section = '\n'.join(section_content)
            current_section = 'greeting'
            remainder = line.split(':', 1)[1].strip() if ':' in line else ''
            section_content = [remainder] if remainder else []
        elif lower_line.startswith('main body:'):
            if section_content and current_section == 'greeting':
                greeting_section = '\n'.join(section_content)
            current_section = 'main body'
            remainder = line.split(':', 1)[1].strip() if ':' in line else ''
            section_content = [remainder] if remainder else []
        elif lower_line.startswith('closing:'):
            if section_content and current_section == 'main body':
                main_body_section = '\n'.join(section_content)
            current_section = 'closing'
            remainder = line.split(':', 1)[1].strip() if ':' in line else ''
            section_content = [remainder] if remainder else []
        elif current_section and line.strip():
            section_content.append(line)
    
    # Finalize last section
    if section_content:
        if current_section == 'greeting':
            greeting_section = '\n'.join(section_content)
        elif current_section == 'main body':
            main_body_section = '\n'.join(section_content)
        elif current_section == 'closing':
            closing_section = '\n'.join(section_content)
    
    # Wrap sections in paragraph tags
    if greeting_section.strip():
        greeting_section = f'<p style="margin:0 0 16px 0;">{greeting_section.strip()}</p>'
    if main_body_section.strip():
        # Split by double newline or keep as-is
        paragraphs = main_body_section.strip().split('\n\n') if '\n\n' in main_body_section else [main_body_section.strip()]
        main_body_section = ''.join(f'<p style="margin:0 0 12px 0;">{p.strip()}</p>' for p in paragraphs if p.strip())
    if closing_section.strip():
        closing_section = f'<p style="margin:0 0 16px 0;">{closing_section.strip()}</p>'

    # Sanitize fragments: strip full-page wrappers or header/footer blocks the model may have included
    def _sanitize_fragment(s: str) -> str:
        if not s:
            return s
        s = re.sub(r'(?is)<\s*head[^>]*>.*?<\s*/\s*head\s*>', '', s)
        s = re.sub(r'(?is)<\s*html[^>]*>', '', s)
        s = re.sub(r'(?is)<\s*/\s*html\s*>', '', s)
        s = re.sub(r'(?is)<\s*body[^>]*>', '', s)
        s = re.sub(r'(?is)<\s*/\s*body\s*>', '', s)
        s = re.sub(r'(?is)<\s*header[^>]*>.*?<\s*/\s*header\s*>', '', s)
        s = re.sub(r'(?is)<\s*h1[^>]*>.*?EduAIHub.*?<\s*/\s*h1\s*>', '', s)
        s = re.sub(r'(?is)<\s*footer[^>]*>.*?<\s*/\s*footer\s*>', '', s)
        s = re.sub(r'(?is)<a[^>]*href=["\"][^"\"]*eduaihub[^"\"]["\"][^>]*>.*?<\s*/\s*a\s*>', '', s)
        s = re.sub(r'\n{3,}', '\n\n', s).strip()
        return s

    greeting_section = _sanitize_fragment(greeting_section)
    main_body_section = _sanitize_fragment(main_body_section)
    closing_section = _sanitize_fragment(closing_section)

    rendered = render_template('email_template_greeting.html', 
                               recipient_name=recipient_name, 
                               ai_hook=hook or subject, 
                               ai_greeting=greeting_section,
                               ai_main_body=main_body_section,
                               ai_closing=closing_section)
    if return_fragments:
        # Return the pieces so callers can re-render the template (avoids nested headers/footers)
        return subject, greeting_section, main_body_section, closing_section
    return subject, rendered
