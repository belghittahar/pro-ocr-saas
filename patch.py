import os
import streamlit as st

# Locate Streamlit's internal index.html file on the server
index_path = os.path.join(os.path.dirname(st.__file__), 'static', 'index.html')

with open(index_path, 'r', encoding='utf-8') as f:
    html = f.read()

adsense_code = '<script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-3336794319990692" crossorigin="anonymous"></script>'

if adsense_code not in html:
    patched_html = html.replace('<head>', f'<head>\n    {adsense_code}')
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(patched_html)
    print("AdSense code injected successfully into the build!")
