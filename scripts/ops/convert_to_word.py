import os
import re

try:
    from markdown_it import MarkdownIt
except ImportError:
    print("markdown-it-py not found.")
    MarkdownIt = None

def protect_math(text):
    math_blocks = []
    
    # Protect block math $$ ... $$
    def repl_block(m):
        math_blocks.append(m.group(0))
        return f"<!--MATHBLOCK{len(math_blocks)-1}-->"
    
    # Protect inline math $ ... $
    def repl_inline(m):
        math_blocks.append(m.group(0))
        return f"<!--MATHINLINE{len(math_blocks)-1}-->"

    text = re.sub(r'\$\$(.*?)\$\$', repl_block, text, flags=re.DOTALL)
    text = re.sub(r'(?<!\$)\$(?!\$)(.*?)(?<!\$)\$(?!\$)', repl_inline, text)
    return text, math_blocks

def restore_math(text, math_blocks):
    for i, block in enumerate(math_blocks):
        text = text.replace(f"<!--MATHBLOCK{i}-->", block)
        text = text.replace(f"&lt;!--MATHBLOCK{i}--&gt;", block)
        text = text.replace(f"<!--MATHINLINE{i}-->", block)
        text = text.replace(f"&lt;!--MATHINLINE{i}--&gt;", block)
    return text

def convert_md_to_html(input_file, output_file):
    if not os.path.exists(input_file):
        print(f"Error: {input_file} not found.")
        return

    with open(input_file, 'r', encoding='utf-8') as f:
        md_content = f.read()

    # Step 1: Protect Math from Markdown parser
    protected_md, math_blocks = protect_math(md_content)

    # Step 2: Render Markdown
    if MarkdownIt:
        md = MarkdownIt("default", {"html": True})
        html_content = md.render(protected_md)
    else:
        html_content = f"<pre>{protected_md}</pre>"

    # Step 3: Restore Math
    html_content = restore_math(html_content, math_blocks)

    # Add MathJax configuration for inline math
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>README Optimization Logic - Word Ready</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; max-width: 900px; margin: 0 auto; padding: 20px; }}
            h1, h2, h3, h4 {{ color: #2c3e50; border-bottom: 1px solid #eee; padding-bottom: 10px; }}
            code {{ background-color: #f4f4f4; padding: 2px 4px; border-radius: 4px; font-family: 'Courier New', Courier, monospace; }}
            pre {{ background-color: #f4f4f4; padding: 15px; border-radius: 5px; overflow-x: auto; }}
            table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
            th {{ background-color: #f8f9fa; }}
        </style>
        <script>
        MathJax = {{
          tex: {{
            inlineMath: [['$', '$'], ['\\\\(', '\\\\)']]
          }},
          svg: {{
            fontCache: 'global'
          }}
        }};
        </script>
        <script type="text/javascript" id="MathJax-script" async
          src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js">
        </script>
    </head>
    <body>
        {html_content}
    </body>
    </html>
    """

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(full_html)
    
    print(f"Directoire de sortie : {os.path.abspath(output_file)}")
    print("Succes : Fichier HTML genere pour Microsoft Word.")

if __name__ == "__main__":
    convert_md_to_html("README_Optimization_Logic.md", "README_Optimization_Logic.html")
