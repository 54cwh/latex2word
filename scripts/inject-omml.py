#!/usr/bin/env python3
import re, subprocess, sys, os, zipfile, tempfile, json

def extract_omml_list(formulas_tex):
    with tempfile.TemporaryDirectory() as tmp:
        tex = os.path.join(tmp, 'f.tex')
        with open(tex, 'w') as f:
            f.write(formulas_tex)
        docx = os.path.join(tmp, 'f.docx')
        subprocess.run(
            ['pandoc', '-f', 'latex', '--to', 'docx', '-o', docx, tex],
            capture_output=True, check=True, timeout=30
        )
        with zipfile.ZipFile(docx) as z:
            xml = z.read('word/document.xml').decode('utf-8')
    return re.findall(r'<m:oMath>.*?</m:oMath>', xml, re.DOTALL)

def simplify_toc_sdt(doc):
    while True:
        sdt_start = doc.find('<w:sdt>')
        if sdt_start == -1:
            break
        content_start = doc.find('<w:sdtContent>', sdt_start, sdt_start + 2000)
        if content_start == -1:
            break
        content_end = doc.find('</w:sdtContent>', content_start)
        sdt_end = doc.find('</w:sdt>', content_end)
        if content_end == -1 or sdt_end == -1:
            break
        sdt_end_tag = sdt_end + len('</w:sdt>')
        inner = doc[content_start + len('<w:sdtContent>'):content_end]
        doc = doc[:sdt_start] + inner + doc[sdt_end_tag:]
    return doc

def set_math_font(settings_xml, font='Cambria Math'):
    if '<w:mathPr>' in settings_xml:
        return settings_xml
    math_pr = f'<w:mathPr><w:mathFont w:val="{font}"/></w:mathPr>'
    ins = settings_xml.find('</w:settings>')
    if ins == -1:
        return settings_xml
    return settings_xml[:ins] + math_pr + settings_xml[ins:]

def inject(our_docx, output_docx, omml_list):
    with zipfile.ZipFile(our_docx, 'r') as z:
        names = z.namelist()
        items = {n: z.read(n) for n in names}

    doc = items['word/document.xml'].decode('utf-8')
    doc = simplify_toc_sdt(doc)

    placeholders = []
    for m in re.finditer(r'__MATH_(\d+)__', doc):
        idx = int(m.group(1))
        ostart = doc.rfind('<m:math', 0, m.start())
        oend = doc.find('</m:math>', m.end())
        has_wrapper = ostart != -1 and oend != -1
        trailing = ''
        if has_wrapper:
            oend += len('</m:math>')
            placeholder_run_end = doc.find('</m:r>', m.end())
            inner_end = doc.rfind('</m:oMath>', m.end(), oend - len('</m:math>'))
            if placeholder_run_end != -1 and inner_end != -1:
                trailing = doc[placeholder_run_end + len('</m:r>'):inner_end].strip()
        else:
            ostart = doc.rfind('<m:oMath', 0, m.start())
            oend = doc.find('</m:oMath>', m.end())
            if ostart == -1 or oend == -1:
                print(f"Warning: MATH_{idx} has no m:math or m:oMath wrapper", file=sys.stderr)
                continue
            oend += len('</m:oMath>')
            placeholder_run_end = doc.find('</m:r>', m.end())
            inner_end = oend - len('</m:oMath>')
            if placeholder_run_end != -1:
                trailing = doc[placeholder_run_end + len('</m:r>'):inner_end].strip()
        placeholders.append((idx, ostart, oend, has_wrapper, trailing))

    placeholders.sort(key=lambda x: -x[1])
    for idx, start, end, has_wrapper, carry in placeholders:
        if idx >= len(omml_list):
            print(f"Warning: MATH_{idx} out of range (pandoc produced {len(omml_list)} OMML)", file=sys.stderr)
            continue
        new_omml = omml_list[idx]
        if carry and len(carry) > 5:
            new_omml = new_omml[:-len('</m:oMath>')] + carry + '</m:oMath>'
        replacement = f'<m:math>{new_omml}</m:math>' if has_wrapper else new_omml
        doc = doc[:start] + replacement + doc[end:]

    items['word/document.xml'] = doc.encode('utf-8')

    if 'word/settings.xml' in items:
        settings = items['word/settings.xml'].decode('utf-8')
        settings = set_math_font(settings)
        items['word/settings.xml'] = settings.encode('utf-8')

    with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, items[n])
    return True

def main():
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('input_docx')
    p.add_argument('output_docx')
    p.add_argument('--formulas-tex')
    p.add_argument('--formulas-json')
    p.add_argument('--omml-json')
    args = p.parse_args()

    omml_list = None
    if args.omml_json:
        with open(args.omml_json) as f:
            omml_list = json.load(f)
    elif args.formulas_tex:
        with open(args.formulas_tex) as f:
            tex = f.read()
        omml_list = extract_omml_list(tex)
        print(f"Extracted {len(omml_list)} OMML formulas from pandoc", file=sys.stderr)
    elif args.formulas_json:
        with open(args.formulas_json) as f:
            formulas = json.load(f)
        lines = [
            '\\documentclass{article}',
            '\\usepackage{amsmath,amssymb}',
            '\\begin{document}',
        ]
        for f in formulas:
            if f['display']:
                lines.append('\\[' + f['latex'] + '\\]')
            else:
                lines.append('\\(' + f['latex'] + '\\)')
        lines.append('\\end{document}')
        omml_list = extract_omml_list('\n'.join(lines))
        print(f"Extracted {len(omml_list)} OMML formulas from pandoc", file=sys.stderr)
    else:
        print("Error: provide --formulas-tex, --formulas-json, or --omml-json", file=sys.stderr)
        sys.exit(1)

    inject(args.input_docx, args.output_docx, omml_list)
    print(f"Injected → {args.output_docx}", file=sys.stderr)

if __name__ == '__main__':
    main()
