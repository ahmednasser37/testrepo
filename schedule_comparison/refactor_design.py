import re

with open('static/style.css', 'r', encoding='utf-8') as f:
    css = f.read()

# Fix the messy :root blocks and syntax errors
css = re.sub(r'/\* Carbon Productive Motion \*/.*?\}', '', css, flags=re.DOTALL)
css = re.sub(r'--motion-duration-moderate-02: 240ms;.*?\}', '', css, flags=re.DOTALL)
css = re.sub(r':root\s*\{\s*\}', '', css)

# Replace variables
replacements = [
    ('--bg:          #f4f4f4;', '--bg:          #f5f5f7;'),
    ('--surface:     #ffffff;', '--surface:     #ffffff;'),
    ('--surface-2:   #e0e0e0;', '--surface-2:   #e5e5ea;'),
    ('--border:      #c6c6c6;', '--border:      #d2d2d7;'),
    ('--border-strong: #8d8d8d;', '--border-strong: #86868b;'),

    ('--text-primary:   #161616;', '--text-primary:   #1d1d1f;'),
    ('--text-secondary: #525252;', '--text-secondary: #86868b;'),
    ('--text-muted:     #a8a8a8;', '--text-muted:     #aeaeb2;'),
    
    ('--brand:        #0f62fe;', '--brand:        #1a73e8;'),
    ('--brand-hover:  #0353e9;', '--brand-hover:  #1557b0;'),
    ('--brand-subtle: #edf5ff;', '--brand-subtle: #e8f0fe;'),
    ('--brand-text:   #0043ce;', '--brand-text:   #174ea6;'),

    ('--shadow-sm: none;', '--shadow-sm: 0 2px 8px rgba(0,0,0,0.04);'),
    ('--shadow:    0 2px 6px 0 rgba(0,0,0,0.30);', '--shadow:    0 4px 16px rgba(0,0,0,0.06), 0 1px 4px rgba(0,0,0,0.04);'),
    ('--shadow:    0 2px 6px 0 rgba(0,0,0,0.15);', '--shadow:    0 4px 16px rgba(0,0,0,0.06), 0 1px 4px rgba(0,0,0,0.04);'),
    ('--shadow-md: 0 4px 8px 0 rgba(0,0,0,0.40);', '--shadow-md: 0 8px 32px rgba(0,0,0,0.08), 0 2px 8px rgba(0,0,0,0.04);'),
    ('--shadow-md: 0 4px 8px 0 rgba(0,0,0,0.20);', '--shadow-md: 0 8px 32px rgba(0,0,0,0.08), 0 2px 8px rgba(0,0,0,0.04);'),

    ('--radius-sm: 0px;', '--radius-sm: 8px;'),
    ('--radius:    0px;', '--radius:    16px;'),
    ('--radius-lg: 0px;', '--radius-lg: 24px;'),
    ('--radius-xl: 0px;', '--radius-xl: 32px;'),

    ('--bg:          #161616;', '--bg:          #000000;'),
    ('--surface:     #262626;', '--surface:     #1c1c1e;'),
    ('--surface-2:   #393939;', '--surface-2:   #2c2c2e;'),
    ('--border:      #525252;', '--border:      #38383a;'),
    
    ('IBM Plex Sans', '-apple-system, BlinkMacSystemFont, "SF Pro Display", "Product Sans", "Roboto", "Segoe UI", sans-serif'),
    ('var(--motion-duration-fast-02) var(--motion-easing-standard)', '0.3s cubic-bezier(0.25, 1, 0.5, 1)'),
    ('var(--motion-duration-moderate-02) var(--motion-easing-expressive)', '0.4s cubic-bezier(0.25, 1, 0.5, 1)'),
    ('var(--motion-duration-fast-01) var(--motion-easing-standard)', '0.2s cubic-bezier(0.25, 1, 0.5, 1)'),
    ('var(--motion-duration-moderate-01) var(--motion-easing-standard)', '0.35s cubic-bezier(0.25, 1, 0.5, 1)'),
    ('border-radius: 0px;', 'border-radius: var(--radius-sm);'),
]

for old, new in replacements:
    css = css.replace(old, new)

with open('static/style.css', 'w', encoding='utf-8') as f:
    f.write(css)

# Remove IBM font imports
for path in ['templates/index.html', 'templates/dashboard.html', 'static/style.css']:
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    content = re.sub(r'<link rel="preconnect" href="https://fonts\.googleapis\.com">.*?IBM\+Plex\+Sans.*?rel="stylesheet">', '', content, flags=re.DOTALL)
    content = re.sub(r'@import url\(\'https://fonts\.googleapis\.com.*?IBM\+Plex\+Sans.*?\'\);', '', content)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
