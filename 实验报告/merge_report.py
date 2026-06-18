#!/usr/bin/env python3
"""
将模板封面（含校标）与现有报告正文合并，生成最终报告。

策略：ZIP级别手术
1. 从模板中提取封面部分（含合肥工业大学校标）
2. 从现有报告中提取正文内容（一、课题概述 到 六、参考文献）
3. 合并两者的图片资源和关系
4. 生成新的 docx 文件
"""

import zipfile
import shutil
import os
import re
import copy
from lxml import etree

# ============================================================
# CONFIG
# ============================================================
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Find template file
TEMPLATE_PATH = None
for f in os.listdir(REPO_ROOT):
    if f.endswith('.docx') and '模板' in f:
        TEMPLATE_PATH = os.path.join(REPO_ROOT, f)
        break

REPORT_PATH = os.path.join(REPO_ROOT, '实验报告', '实验报告_莫仁鹰_2023212167.docx')
OUTPUT_PATH = os.path.join(REPO_ROOT, '实验报告', '实验报告_莫仁鹰_2023212167_v2.docx')

NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v': 'urn:schemas-microsoft-com:vml',
    'o': 'urn:schemas-microsoft-com:office:office',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
}

REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
RELS_NSMAP = {None: REL_NS}

print(f"Template: {TEMPLATE_PATH}")
print(f"Report:   {REPORT_PATH}")
print(f"Output:   {OUTPUT_PATH}")

# ============================================================
# STEP 1: Read both docx files
# ============================================================

def read_docx_xml(docx_path, inner_path):
    """Read and parse an XML file from inside a docx ZIP."""
    with zipfile.ZipFile(docx_path, 'r') as z:
        return etree.fromstring(z.read(inner_path))

def list_media(docx_path):
    """List all media files in a docx."""
    with zipfile.ZipFile(docx_path, 'r') as z:
        return [f for f in z.namelist() if f.startswith('word/media/')]

def read_rels(docx_path):
    """Read document.xml.rels and return parsed XML."""
    with zipfile.ZipFile(docx_path, 'r') as z:
        return etree.fromstring(z.read('word/_rels/document.xml.rels'))

# Parse document XMLs
template_doc = read_docx_xml(TEMPLATE_PATH, 'word/document.xml')
report_doc = read_docx_xml(REPORT_PATH, 'word/document.xml')

template_body = template_doc.find('.//w:body', NSMAP)
report_body = report_doc.find('.//w:body', NSMAP)

# ============================================================
# STEP 2: Identify cover page boundary in template
# ============================================================

def get_element_text(elem):
    """Get all text content of an XML element."""
    return ''.join(elem.itertext()).strip()

# Template structure:
# Elements [0]-[12]: Cover page (including section break paragraph)
# Element [12] has the section break (<w:sectPr> in <w:pPr>)
# Elements [13]-[60]: Placeholder content
# Element [61]: Final <w:sectPr> (direct child of body)

template_children = list(template_body)

# Find the section break paragraph (contains sectPr in pPr)
cover_end_idx = None
for i, child in enumerate(template_children):
    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
    if tag == 'p':
        pPr = child.find('w:pPr', NSMAP)
        if pPr is not None:
            sectPr = pPr.find('w:sectPr', NSMAP)
            if sectPr is not None:
                cover_end_idx = i
                print(f"Template: Section break at element [{i}]")
                break

if cover_end_idx is None:
    # Fallback: use the paragraph just before "课题概述" content
    for i, child in enumerate(template_children):
        text = get_element_text(child)
        if '课题概述' in text or '课题任务' in text:
            cover_end_idx = i - 1
            print(f"Template: Content starts at element [{i}], cover ends at [{cover_end_idx}]")
            break

# The final sectPr (last element in body, direct child)
final_sectPr_idx = len(template_children) - 1
assert template_children[final_sectPr_idx].tag.endswith('}sectPr') or \
       template_children[final_sectPr_idx].tag == 'sectPr', \
       f"Last element is not sectPr: {template_children[final_sectPr_idx].tag}"

print(f"Template: Cover elements [0..{cover_end_idx}], final sectPr at [{final_sectPr_idx}]")

# ============================================================
# STEP 3: Identify content boundary in existing report
# ============================================================

report_children = list(report_body)

# Find "一、课题概述" in the existing report
content_start_idx = None
for i, child in enumerate(report_children):
    text = get_element_text(child)
    if '课题概述' in text and len(text) < 20:
        content_start_idx = i
        print(f"Report: Content starts at element [{i}] ('{text[:30]}')")
        break

# Find the final sectPr in report
report_final_sectPr_idx = len(report_children) - 1
print(f"Report: {len(report_children)} elements, content from [{content_start_idx}] to [{report_final_sectPr_idx}]")

# ============================================================
# STEP 4: Collect relationship IDs from both documents
# ============================================================

template_rels = read_rels(TEMPLATE_PATH)
report_rels = read_rels(REPORT_PATH)

def get_rels_dict(rels_xml):
    """Parse relationships into a dict: rId -> (type, target)"""
    result = {}
    for rel in rels_xml:
        rid = rel.get('Id')
        rtype = rel.get('Type')
        target = rel.get('Target')
        result[rid] = (rtype, target)
    return result

template_rels_dict = get_rels_dict(template_rels)
report_rels_dict = get_rels_dict(report_rels)

print(f"\nTemplate relationships: {len(template_rels_dict)}")
print(f"Report relationships: {len(report_rels_dict)}")

# Find the highest rId number in template
def max_rid(rels_dict):
    max_n = 0
    for rid in rels_dict:
        m = re.match(r'rId(\d+)', rid)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n

template_max_rid = max_rid(template_rels_dict)
print(f"Template max rId: rId{template_max_rid}")

# ============================================================
# STEP 5: Collect rIds used in report content
# ============================================================

def find_all_rids(elements):
    """Find all r:id and r:embed references in a list of XML elements."""
    rids = set()
    r_ns = NSMAP['r']
    for elem in elements:
        for attr_name in [f'{{{r_ns}}}id', f'{{{r_ns}}}embed', f'{{{r_ns}}}link']:
            val = elem.get(attr_name)
            if val:
                rids.add(val)
        for child in elem.iter():
            for attr_name in [f'{{{r_ns}}}id', f'{{{r_ns}}}embed', f'{{{r_ns}}}link']:
                val = child.get(attr_name)
                if val:
                    rids.add(val)
    return rids

# Content elements from report (everything from content_start to before final sectPr)
report_content_elements = report_children[content_start_idx:report_final_sectPr_idx]
report_content_rids = find_all_rids(report_content_elements)
print(f"\nReport content uses {len(report_content_rids)} relationship IDs: {sorted(report_content_rids)}")

# ============================================================
# STEP 6: Build rId remapping for report content
# ============================================================

# We need to remap report rIds to avoid conflicts with template rIds
rid_remap = {}
next_rid = template_max_rid + 1

for old_rid in sorted(report_content_rids):
    if old_rid in template_rels_dict:
        # Conflict! Need to remap
        new_rid = f'rId{next_rid}'
        rid_remap[old_rid] = new_rid
        next_rid += 1
    else:
        # No conflict, but let's remap anyway to be safe
        new_rid = f'rId{next_rid}'
        rid_remap[old_rid] = new_rid
        next_rid += 1

print(f"\nrId remapping: {len(rid_remap)} entries")
for old, new in sorted(rid_remap.items()):
    if old in report_rels_dict:
        _, target = report_rels_dict[old]
        print(f"  {old} -> {new} (target: {target})")

# ============================================================
# STEP 7: Build media file remapping
# ============================================================

template_media = list_media(TEMPLATE_PATH)
report_media = list_media(REPORT_PATH)

# Find max image number in template
template_max_img = 0
for m in template_media:
    match = re.search(r'image(\d+)', m)
    if match:
        template_max_img = max(template_max_img, int(match.group(1)))

print(f"\nTemplate has {len(template_media)} media files, max image{template_max_img}")
print(f"Report has {len(report_media)} media files")

# Remap report media filenames
media_remap = {}
next_img = template_max_img + 1
for media_path in report_media:
    fname = os.path.basename(media_path)
    ext = os.path.splitext(fname)[1]
    new_fname = f'image{next_img}{ext}'
    new_path = f'word/media/{new_fname}'
    media_remap[media_path] = new_path
    next_img += 1

# ============================================================
# STEP 8: Apply rId remapping to report content elements
# ============================================================

def remap_rids_in_elements(elements, rid_remap):
    """Replace all r:id/r:embed references in elements according to mapping."""
    r_ns = NSMAP['r']
    attr_names = [f'{{{r_ns}}}id', f'{{{r_ns}}}embed', f'{{{r_ns}}}link']
    
    for elem in elements:
        for attr_name in attr_names:
            val = elem.get(attr_name)
            if val and val in rid_remap:
                elem.set(attr_name, rid_remap[val])
        for child in elem.iter():
            for attr_name in attr_names:
                val = child.get(attr_name)
                if val and val in rid_remap:
                    child.set(attr_name, rid_remap[val])

# Deep copy report content elements before modifying
report_content_copy = [copy.deepcopy(elem) for elem in report_content_elements]
remap_rids_in_elements(report_content_copy, rid_remap)

# ============================================================
# STEP 9: Fill in the cover page table
# ============================================================

STUDENT_NAME = "莫仁鹰"
STUDENT_ID = "2023212167"
COLLEGE = "计算机与信息学院"
MAJOR_CLASS = "物联网23级-1班"
ADVISOR = "【指导教师姓名】"
TOPIC = "基于RK3566的智能无人飞行器视觉识别系统设计"
DATE_STR = "2026年6月17日"

# Find the table in cover page
cover_elements = template_children[:cover_end_idx + 1]
cover_table = None
for elem in cover_elements:
    if elem.tag.endswith('}tbl') or elem.tag == 'tbl':
        cover_table = elem
        break

if cover_table is not None:
    w_ns = NSMAP['w']
    rows = cover_table.findall(f'{{{w_ns}}}tr')
    
    info_values = [
        COLLEGE,                              # Row 0, Col 1
        MAJOR_CLASS,                          # Row 1, Col 1
        f"{STUDENT_NAME}  {STUDENT_ID}",      # Row 2, Col 1
        ADVISOR,                              # Row 3, Col 1
        TOPIC,                                # Row 4, Col 1
    ]
    
    for row_idx, value in enumerate(info_values):
        if row_idx < len(rows):
            cells = rows[row_idx].findall(f'{{{w_ns}}}tc')
            if len(cells) >= 2:
                # Clear existing content in cell 1
                cell = cells[1]
                for p in cell.findall(f'{{{w_ns}}}p'):
                    # Clear all runs
                    for r in p.findall(f'{{{w_ns}}}r'):
                        for t in r.findall(f'{{{w_ns}}}t'):
                            t.text = value
                            # Only set text in first run, clear others
                        break
                    # Simpler: remove all runs and add one
                    pPr = p.find(f'{{{w_ns}}}pPr')
                    for r in list(p.findall(f'{{{w_ns}}}r')):
                        p.remove(r)
                    # Add a new run with the value
                    new_run = etree.SubElement(p, f'{{{w_ns}}}r')
                    rPr = etree.SubElement(new_run, f'{{{w_ns}}}rPr')
                    rFonts = etree.SubElement(rPr, f'{{{w_ns}}}rFonts')
                    rFonts.set(f'{{{w_ns}}}eastAsia', '宋体')
                    sz = etree.SubElement(rPr, f'{{{w_ns}}}sz')
                    sz.set(f'{{{w_ns}}}val', '28')  # 14pt = 28 half-points
                    new_t = etree.SubElement(new_run, f'{{{w_ns}}}t')
                    new_t.text = value
                    new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    
    # Row 5: Date (both cells)
    if len(rows) > 5:
        cells = rows[5].findall(f'{{{w_ns}}}tc')
        for cell in cells:
            for p in cell.findall(f'{{{w_ns}}}p'):
                pPr = p.find(f'{{{w_ns}}}pPr')
                for r in list(p.findall(f'{{{w_ns}}}r')):
                    p.remove(r)
                new_run = etree.SubElement(p, f'{{{w_ns}}}r')
                rPr = etree.SubElement(new_run, f'{{{w_ns}}}rPr')
                rFonts = etree.SubElement(rPr, f'{{{w_ns}}}rFonts')
                rFonts.set(f'{{{w_ns}}}eastAsia', '宋体')
                sz = etree.SubElement(rPr, f'{{{w_ns}}}sz')
                sz.set(f'{{{w_ns}}}val', '28')
                new_t = etree.SubElement(new_run, f'{{{w_ns}}}t')
                new_t.text = DATE_STR
                new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    
    print("\nCover table filled with student info")

# ============================================================
# STEP 10: Assemble the new document body
# ============================================================

# Clear the template body
for child in list(template_body):
    template_body.remove(child)

w_ns = NSMAP['w']

# Add cover page elements (from template), EXCLUDING the section-break paragraph
# Element [cover_end_idx] has the sectPr embedded. We need the sectPr but not
# the placeholder text. So add elements [0..cover_end_idx-1], then create a
# clean section-break paragraph.
for i in range(cover_end_idx):
    template_body.append(cover_elements[i])

# Create a clean section-break paragraph (preserves cover/content section separation)
sect_break_para = cover_elements[cover_end_idx]
sect_pr_in_ppr = sect_break_para.find(f'{{{w_ns}}}pPr/{{{w_ns}}}sectPr')
if sect_pr_in_ppr is not None:
    # Create a new empty paragraph with just the sectPr
    new_sect_p = etree.SubElement(template_body, f'{{{w_ns}}}p')
    new_pPr = etree.SubElement(new_sect_p, f'{{{w_ns}}}pPr')
    new_pPr.append(copy.deepcopy(sect_pr_in_ppr))
    print("Created clean section break paragraph")
else:
    # No sectPr found, just skip the element
    print("WARNING: No sectPr in cover_end element, skipping")

# Add report content directly (the section break handles the page transition)
for elem in report_content_copy:
    template_body.append(elem)

# Add final sectPr from report (preserves content section formatting)
report_final_sectPr = report_children[report_final_sectPr_idx]
if report_final_sectPr.tag.endswith('}sectPr'):
    template_body.append(copy.deepcopy(report_final_sectPr))
else:
    template_body.append(template_children[final_sectPr_idx])

print(f"New body has {len(template_body)} elements")

# ============================================================
# STEP 11: Build new relationships XML
# ============================================================

# Start with all template relationships
new_rels = copy.deepcopy(template_rels)

# Add remapped report relationships (only those used in content)
IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

for old_rid, new_rid in rid_remap.items():
    if old_rid in report_rels_dict:
        rtype, target = report_rels_dict[old_rid]
        
        # Remap media target if it's an image
        new_target = target
        if target.startswith('media/'):
            full_path = f'word/{target}'
            if full_path in media_remap:
                new_target = 'media/' + os.path.basename(media_remap[full_path])
        
        # Create new relationship element
        new_rel = etree.SubElement(new_rels, 'Relationship')
        new_rel.set('Id', new_rid)
        new_rel.set('Type', rtype)
        new_rel.set('Target', new_target)

print(f"New relationships: {len(new_rels)} entries")

# ============================================================
# STEP 12: Assemble the final docx ZIP
# ============================================================

# Start with a copy of the template
shutil.copy2(TEMPLATE_PATH, OUTPUT_PATH)

# Now modify the copy
with zipfile.ZipFile(TEMPLATE_PATH, 'r') as template_zip:
    with zipfile.ZipFile(REPORT_PATH, 'r') as report_zip:
        # Create new ZIP
        import tempfile
        tmp_path = OUTPUT_PATH + '.tmp'
        
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as out_zip:
            # Copy everything from template except document.xml and rels
            for item in template_zip.namelist():
                if item == 'word/document.xml':
                    continue
                if item == 'word/_rels/document.xml.rels':
                    continue
                if item == '[Content_Types].xml':
                    continue
                out_zip.writestr(item, template_zip.read(item))
            
            # Write new document.xml
            doc_xml = etree.tostring(template_doc, xml_declaration=True, 
                                      encoding='UTF-8', standalone=True)
            out_zip.writestr('word/document.xml', doc_xml)
            
            # Write new relationships
            rels_xml = etree.tostring(new_rels, xml_declaration=True,
                                       encoding='UTF-8', standalone=True)
            out_zip.writestr('word/_rels/document.xml.rels', rels_xml)
            
            # Copy report media files (with remapped names)
            for old_path, new_path in media_remap.items():
                try:
                    data = report_zip.read(old_path)
                    out_zip.writestr(new_path, data)
                except KeyError:
                    print(f"  WARNING: Media file not found: {old_path}")
            
            # Update [Content_Types].xml to include new media types
            content_types = etree.fromstring(template_zip.read('[Content_Types].xml'))
            ct_ns = content_types.nsmap.get(None, '')
            
            # Ensure all image extensions are registered
            existing_extensions = set()
            for default in content_types.findall(f'{{{ct_ns}}}Default'):
                existing_extensions.add(default.get('Extension', '').lower())
            
            ext_map = {
                'png': 'image/png',
                'jpeg': 'image/jpeg',
                'jpg': 'image/jpeg',
                'gif': 'image/gif',
                'bmp': 'image/bmp',
                'wmf': 'image/x-wmf',
                'emf': 'image/x-emf',
            }
            
            for media_file in media_remap.values():
                ext = os.path.splitext(media_file)[1].lstrip('.').lower()
                if ext not in existing_extensions and ext in ext_map:
                    new_default = etree.SubElement(content_types, f'{{{ct_ns}}}Default')
                    new_default.set('Extension', ext)
                    new_default.set('ContentType', ext_map[ext])
                    existing_extensions.add(ext)
            
            ct_xml = etree.tostring(content_types, xml_declaration=True,
                                     encoding='UTF-8', standalone=True)
            out_zip.writestr('[Content_Types].xml', ct_xml)

# Replace output with tmp
os.replace(tmp_path, OUTPUT_PATH)

print(f"\nReport assembled: {OUTPUT_PATH}")
print(f"  File size: {os.path.getsize(OUTPUT_PATH)} bytes")

# ============================================================
# STEP 13: Add watermarks to all report images
# ============================================================
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
import io

WATERMARK_TEXT = f"{datetime.now().strftime('%Y-%m-%d %H:%M')}  {STUDENT_ID}  {STUDENT_NAME}"

# Template images to SKIP (school logo, template examples)
template_media_basenames = set(os.path.basename(m) for m in template_media)
print(f"\nSkipping template images: {template_media_basenames}")
print(f"Watermark text: {WATERMARK_TEXT}")

def add_watermark(img_data, filename):
    """Add watermark to image data, return watermarked bytes."""
    ext = os.path.splitext(filename)[1].lower()
    if ext in ('.wmf', '.emf'):
        return img_data  # Can't watermark vector formats
    
    try:
        img = Image.open(io.BytesIO(img_data))
        if img.mode == 'RGBA':
            # Convert to RGB for JPEG output, keep RGBA for PNG
            pass
        elif img.mode != 'RGB':
            img = img.convert('RGB')
        
        draw = ImageDraw.Draw(img)
        
        # Dynamic font size based on image dimensions
        font_size = max(12, min(img.width, img.height) // 30)
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", font_size)
        except (OSError, IOError):
            try:
                font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", font_size)
            except (OSError, IOError):
                font = ImageFont.load_default()
        
        # Position: bottom-right corner with padding
        bbox = draw.textbbox((0, 0), WATERMARK_TEXT, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        padding = 10
        x = img.width - text_w - padding
        y = img.height - text_h - padding
        
        # Semi-transparent background
        bg_rect = [x - 5, y - 3, img.width - padding + 5, img.height - padding + 3]
        if img.mode == 'RGBA':
            overlay = Image.new('RGBA', img.size, (0, 0, 0, 0))
            overlay_draw = ImageDraw.Draw(overlay)
            overlay_draw.rectangle(bg_rect, fill=(0, 0, 0, 100))
            overlay_draw.text((x, y), WATERMARK_TEXT, fill=(255, 255, 255, 200), font=font)
            img = Image.alpha_composite(img, overlay)
        else:
            # For RGB, just draw directly with some opacity simulation
            draw.rectangle(bg_rect, fill=(0, 0, 0))
            draw.text((x, y), WATERMARK_TEXT, fill=(255, 255, 255), font=font)
        
        # Save back
        output = io.BytesIO()
        if ext in ('.jpg', '.jpeg'):
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            img.save(output, 'JPEG', quality=95)
        else:
            img.save(output, 'PNG')
        
        return output.getvalue()
    except Exception as e:
        print(f"  WARNING: Failed to watermark {filename}: {e}")
        return img_data

# Process the output docx
watermarked_count = 0
with zipfile.ZipFile(OUTPUT_PATH, 'r') as zin:
    tmp_wm_path = OUTPUT_PATH + '.wm.tmp'
    with zipfile.ZipFile(tmp_wm_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            
            if item.startswith('word/media/'):
                basename = os.path.basename(item)
                # Skip template's own images
                if basename not in template_media_basenames:
                    data = add_watermark(data, basename)
                    watermarked_count += 1
                    print(f"  Watermarked: {basename}")
            
            zout.writestr(item, data)

os.replace(tmp_wm_path, OUTPUT_PATH)
print(f"\n✓ Watermarked {watermarked_count} images")
print(f"✓ Final report: {OUTPUT_PATH}")
print(f"  File size: {os.path.getsize(OUTPUT_PATH)} bytes")
print(f"  ({os.path.getsize(OUTPUT_PATH) / 1024 / 1024:.1f} MB)")
