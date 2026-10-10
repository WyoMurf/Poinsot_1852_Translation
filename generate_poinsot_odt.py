import uno, re, os, subprocess, time
from com.sun.star.awt import Size
from com.sun.star.style import LineSpacing, LineSpacingMode

CENTER = uno.Enum('com.sun.star.style.ParagraphAdjust', 'CENTER')
LEFT = uno.Enum('com.sun.star.style.ParagraphAdjust', 'LEFT')
ITALIC = uno.Enum('com.sun.star.awt.FontSlant', 'ITALIC')
NONE = uno.Enum('com.sun.star.awt.FontSlant', 'NONE')
AS_CHAR = uno.Enum('com.sun.star.text.TextContentAnchorType', 'AS_CHARACTER')

PORT = 2049
soffice_cmd = [
    '/opt/libreoffice26.2/program/soffice',
    '--headless',
    f'--accept=socket,host=127.0.0.1,port={PORT};urp;'
]

print("Launching headless LibreOffice...")
proc = subprocess.Popen(soffice_cmd)
time.sleep(2.5)

try:
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        'com.sun.star.bridge.UnoUrlResolver', localContext)
    ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={PORT};urp;StarOffice.ComponentContext')
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)

    print("Connected to LibreOffice UNO bridge!")
    doc = desktop.loadComponentFromURL('private:factory/swriter', '_blank', 0, ())
    doc.lockControllers()

    text = doc.getText()
    cursor = text.createTextCursor()

    # Set default paragraph style properties
    ls_prop = LineSpacing(Mode=LineSpacingMode.PROP, Height=115)

    md_file = '/home/murf/Poinsot_Translation/Poinsot_Translation.md'
    with open(md_file, 'r', encoding='utf-8') as f:
        content = f.read()

    paragraphs = content.split('\n\n')
    total = len(paragraphs)
    print(f"Total paragraphs to process: {total}")

    MATH_CLSID = '078B7ABA-54FC-457F-8551-6147E776A997'
    MATH_OPS = ['=', '+', '-', '*', '/', 'over', 'sqrt', 'int', '∫', '^', '_', 'sum', 'sin', 'cos', 'cdot', 'times', '+-', '<', '>', '{', '}']

    def is_math_formula(expr):
        return any(op in expr for op in MATH_OPS)

    def insert_formula_object(cursor, expr, is_display=False):
        expr = expr.strip()
        if not expr:
            return
        
        # Scale base font size so fractions match 12pt body text
        if is_display:
            target_sz = 15  # 15pt for display equations
        elif 'over' in expr:
            target_sz = 14  # 14pt for inline fractions
        else:
            target_sz = 13  # 13pt for standard inline formulas
        
        formula_str = f"size {target_sz} {{ {expr} }}"
        try:
            f_obj = doc.createInstance('com.sun.star.text.TextEmbeddedObject')
            f_obj.CLSID = MATH_CLSID
            f_obj.AnchorType = AS_CHAR
            text.insertTextContent(cursor, f_obj, False)
            f_obj.Model.Formula = formula_str
            vis = f_obj.Model.getVisualAreaSize(1)
            if vis.Width > 0 and vis.Height > 0:
                f_obj.setSize(vis)
        except Exception:
            cursor.CharPosture = ITALIC
            text.insertString(cursor, expr, False)
            cursor.CharPosture = NONE

    def insert_image(cursor, rel_path, caption=""):
        full_path = os.path.abspath(os.path.join('/home/murf/Poinsot_Translation', rel_path))
        if not os.path.exists(full_path):
            print(f"Warning: Image not found: {full_path}")
            return
        g_obj = doc.createInstance('com.sun.star.text.TextGraphicObject')
        g_obj.GraphicURL = uno.systemPathToFileUrl(full_path)
        g_obj.AnchorType = AS_CHAR
        text.insertTextContent(cursor, g_obj, False)
        act = g_obj.ActualSize
        if act.Width > 0:
            max_w = 16000 if 'plate_' in rel_path else 11000
            target_w = min(act.Width, max_w)
            aspect = act.Height / act.Width
            target_h = int(target_w * aspect)
            g_obj.setSize(Size(target_w, target_h))

    t0 = time.time()
    for idx, p in enumerate(paragraphs, 1):
        p = p.strip()
        if not p:
            continue

        if idx % 200 == 0 or idx == total:
            print(f"  Processed {idx}/{total} paragraphs ({time.time() - t0:.1f}s)...")

        # Document Title
        if p.startswith('# '):
            cursor.ParaStyleName = 'Title'
            cursor.ParaAdjust = CENTER
            text.insertString(cursor, p[2:].strip() + '\n', False)
            cursor.ParaAdjust = LEFT
            continue
        elif p.startswith('By M. Poinsot'):
            cursor.ParaStyleName = 'Subtitle'
            cursor.ParaAdjust = CENTER
            text.insertString(cursor, p.strip() + '\n\n', False)
            cursor.ParaAdjust = LEFT
            continue
        # Headings
        elif p.startswith('## '):
            cursor.ParaStyleName = 'Heading 1'
            text.insertString(cursor, p[3:].strip() + '\n', False)
            continue
        elif p.startswith('### '):
            heading_text = p[4:].strip()
            if heading_text.startswith('Page '):
                cursor.ParaStyleName = 'Text body'
                cursor.ParaAdjust = CENTER
                cursor.CharColor = 0x888888
                cursor.CharHeight = 9.0
                text.insertString(cursor, f"— [Page {heading_text[5:]}] —\n", False)
                cursor.CharHeight = 12.0
                cursor.CharColor = 0x000000
                cursor.ParaAdjust = LEFT
            else:
                cursor.ParaStyleName = 'Heading 2'
                text.insertString(cursor, heading_text + '\n', False)
            continue
        elif p.startswith('#### '):
            cursor.ParaStyleName = 'Heading 3'
            text.insertString(cursor, p[5:] + '\n', False)
            continue
        elif p.startswith('---'):
            continue

        # Standalone Image
        img_match = re.match(r'!\[(.*?)\]\((.*?)\)', p)
        if img_match:
            caption = img_match.group(1)
            rel_path = img_match.group(2)
            cursor.ParaStyleName = 'Text body'
            cursor.ParaAdjust = CENTER
            insert_image(cursor, rel_path, caption)
            text.insertString(cursor, '\n', False)
            cursor.ParaAdjust = LEFT
            continue

        # Display formula paragraph
        lines_in_p = [l.strip() for l in p.split('\n') if l.strip()]
        is_display_para = all(
            bool(re.match(r'^(`[^`]+`(\s*,\s*`[^`]+`)*|\s*`[^`]+`\s*)+$', l))
            for l in lines_in_p
        )

        if is_display_para:
            cursor.ParaStyleName = 'Text body'
            cursor.ParaAdjust = CENTER
            cursor.ParaTopMargin = 300    # 3mm padding above display equations
            cursor.ParaBottomMargin = 300 # 3mm padding below display equations
            for l in lines_in_p:
                formula_blocks = re.findall(r'`([^`]+)`', l)
                for f_idx, fb in enumerate(formula_blocks):
                    if f_idx > 0:
                        text.insertString(cursor, ',   ', False)
                    insert_formula_object(cursor, fb, is_display=True)
                text.insertString(cursor, '\n', False)
            cursor.ParaTopMargin = 0
            cursor.ParaBottomMargin = 0
            cursor.ParaAdjust = LEFT
            continue

        # Regular text paragraph with inline math
        cursor.ParaStyleName = 'Text body'
        cursor.ParaLineSpacing = ls_prop
        parts = re.split(r'(`[^`]+`)', p)
        for part in parts:
            if part.startswith('`') and part.endswith('`'):
                expr = part[1:-1].strip()
                if is_math_formula(expr):
                    insert_formula_object(cursor, expr, is_display=False)
                else:
                    cursor.CharPosture = ITALIC
                    text.insertString(cursor, expr, False)
                    cursor.CharPosture = NONE
            else:
                clean_part = re.sub(r'(?<!\n)\n(?!\n)', ' ', part)
                text.insertString(cursor, clean_part, False)
        text.insertString(cursor, '\n\n', False)

    print(f"All paragraphs processed in {time.time()-t0:.1f}s. Unlocking controllers...")
    doc.unlockControllers()

    out_file = '/home/murf/Poinsot_Translation/Poinsot_Translation.odt'
    tmp_out = '/home/murf/Poinsot_Translation/Poinsot_Translation_new.odt'
    if os.path.exists(tmp_out):
        os.remove(tmp_out)
    print(f"Saving document to {tmp_out}...")
    t_save = time.time()
    out_url = uno.systemPathToFileUrl(os.path.abspath(tmp_out))
    doc.storeToURL(out_url, ())
    doc.close(True)
    os.replace(tmp_out, out_file)
    print(f"Saved in {time.time()-t_save:.1f}s. Document size: {os.path.getsize(out_file):,} bytes.")

finally:
    proc.terminate()
    time.sleep(1)
