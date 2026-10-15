"""Browser-local HWP/HWPX reader and independent template PPTX builder.

No manuscript is sent to a server. This module also runs in CPython for tests.
"""
import bisect
import copy
import io
import json
import posixpath
import re
import struct
import sys
import tempfile
import zipfile
import zlib
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).parent / 'reference' / 'scripts'))
sys.path.insert(0, str(Path(__file__).parent / 'vendor'))
import build as template

MAX_UNPACKED = 120 * 1024 * 1024
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
ET.register_namespace('p', P)
ET.register_namespace('r', R)


def xml(data):
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise ValueError('외부 참조가 들어 있는 문서는 지원하지 않습니다.')
    return ET.fromstring(data)


def checked_zip(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    if len(z.infolist()) > 12000 or sum(f.file_size for f in z.infolist()) > MAX_UNPACKED:
        raise ValueError('문서의 압축 해제 크기가 너무 큽니다. 사진을 줄여 주세요.')
    for f in z.infolist():
        if f.filename.startswith('/') or '..' in Path(f.filename).parts or '\\' in f.filename:
            raise ValueError('올바르지 않은 파일 경로입니다.')
    return z


def local(tag):
    return tag.rsplit('}', 1)[-1]


def classify(color):
    try:
        r, g, b = bytes.fromhex(color.lstrip('#')[:6])
    except (ValueError, TypeError):
        return 'BLACK'
    if max(r, g, b) < 80:
        return 'BLACK'
    if b > r + 60 and b > g + 40:
        return 'BLUE'
    if g > r + 50 and g > b + 30:
        return 'GREEN'
    if r > g + 80 and r > b + 80:
        return 'RED'
    return color.upper()


def merge(runs):
    out = []
    for color, text in runs:
        if not text:
            continue
        if out and out[-1][0] == color:
            out[-1][1] += text
        else:
            out.append([color, text])
    return out


def read_hwpx(data):
    z = checked_zip(data)
    header = xml(z.read('Contents/header.xml'))
    colors = {e.get('id'): classify(e.get('textColor', '#000000'))
              for e in header.iter() if local(e.tag) == 'charPr'}
    paragraphs = []
    sections = sorted((n for n in z.namelist() if re.fullmatch(r'Contents/section\d+\.xml', n)),
                      key=lambda n: int(re.search(r'\d+', n).group()))
    for name in sections:
        for p in xml(z.read(name)).iter():
            if local(p.tag) != 'p':
                continue
            runs = []
            for run in p:
                if local(run.tag) != 'run':
                    continue
                text = ''
                for t in run:
                    if local(t.tag) != 't':
                        continue
                    text += t.text or ''
                    for c in t:
                        if local(c.tag) in ('lineBreak', 'br'):
                            text += '\n'
                        elif local(c.tag) == 'tab':
                            text += ' '
                        text += c.tail or ''
                runs.append((colors.get(run.get('charPrIDRef'), 'BLACK'), text))
            if any(t.strip() for _, t in runs):
                paragraphs.append(merge(runs))
    pictures = sum(1 for name in sections for e in xml(z.read(name)).iter() if local(e.tag) == 'pic')
    warnings = [f'원고에 그림 {pictures}개가 있습니다. 그림 자동 배치는 지원하지 않으므로 PPT에서 추가해 주세요.'] if pictures else []
    return paragraphs, warnings


def records(data):
    pos = 0
    while pos + 4 <= len(data):
        h, = struct.unpack_from('<I', data, pos)
        pos += 4
        size = h >> 20
        if size == 4095:
            if pos + 4 > len(data):
                raise ValueError('한글 문서의 기록이 잘렸습니다.')
            size, = struct.unpack_from('<I', data, pos)
            pos += 4
        if pos + size > len(data):
            raise ValueError('한글 문서의 기록 길이가 맞지 않습니다.')
        yield h & 1023, data[pos:pos + size]
        pos += size


def read_hwp(data):
    import olefile
    with olefile.OleFileIO(io.BytesIO(data)) as ole:
        header = ole.openstream('FileHeader').read()
        if not header.startswith(b'HWP Document File'):
            raise ValueError('HWP 5 형식이 아닙니다. 한글에서 HWPX로 저장해 주세요.')
        flags, = struct.unpack_from('<I', header, 36)
        if flags & 6:
            raise ValueError('암호 또는 배포용 한글 문서는 HWPX로 다시 저장해 주세요.')
        def stream(name):
            raw = ole.openstream(name).read()
            if flags & 1:
                dec = zlib.decompressobj(-15)
                raw = dec.decompress(raw, MAX_UNPACKED + 1)
                if len(raw) > MAX_UNPACKED or dec.unconsumed_tail:
                    raise ValueError('한글 문서의 압축 해제 크기가 너무 큽니다.')
            return raw
        colors = []
        for tag, raw in records(stream('DocInfo')):
            if tag == 21:
                if len(raw) < 56:
                    raise ValueError('글자 색상 정보를 읽지 못했습니다.')
                colors.append(classify('#' + raw[52:55].hex()))
        paragraphs = []
        sections = sorted((p for p in ole.listdir() if len(p) == 2 and p[0] == 'BodyText' and re.fullmatch(r'Section\d+', p[1])),
                          key=lambda p: int(p[1][7:]))
        for sec in sections:
            text = None
            spans = []
            def flush():
                if text is None:
                    return
                positions = [p for p, _ in spans]
                runs = []
                i = 0
                while i < len(text) // 2:
                    ch, = struct.unpack_from('<H', text, i * 2)
                    at = max(0, bisect.bisect_right(positions, i) - 1)
                    cid = spans[at][1] if spans else 0
                    color = colors[cid] if cid < len(colors) else 'BLACK'
                    if ch >= 32:
                        length = 2 if 0xD800 <= ch <= 0xDBFF else 1
                        runs.append((color, text[i * 2:(i + length) * 2].decode('utf-16le', errors='replace')))
                        i += length
                        continue
                    if ch in (9, 10):
                        runs.append((color, ' ' if ch == 9 else '\n'))
                    i += 8 if ch in (1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23) else 1
                if any(t.strip() for _, t in runs):
                    paragraphs.append(merge(runs))
            for tag, raw in records(stream(sec)):
                if tag == 66:
                    flush()
                    text, spans = None, []
                elif tag == 67:
                    text = raw
                elif tag == 68:
                    spans = list(struct.iter_unpack('<II', raw))
            flush()
        pictures = [p for p in ole.listdir() if p[0] == 'BinData']
    warnings = ['원고에 그림·개체가 있습니다. 그림 자동 배치는 지원하지 않으므로 PPT에서 확인해 주세요.'] if pictures else []
    return paragraphs, warnings


REF = re.compile(r'[（(<〈]?([가-힣]{1,12}\s*\d+\s*(?:장|:)\s*\d+(?:\s*[~\-–,]\s*\d+)*(?:절)?)[）)>〉]?')


def wrap_marked(text, width):
    # Prefer word boundaries; preserve inline color across every new line.
    chars, active = [], None
    for token in re.findall(r'</?[rb]>|.', text, re.S):
        if token in ('<r>', '<b>'):
            active = token[1]
        elif token in ('</r>', '</b>'):
            active = None
        else:
            chars.append((token, active))
    paragraphs, current = [], []
    for pair in chars:
        if pair[0] == '\n':
            paragraphs.append(current)
            current = []
        else:
            current.append(pair)
    paragraphs.append(current)
    lines = []
    for paragraph in paragraphs:
        while paragraph:
            while paragraph and paragraph[0][0].isspace():
                paragraph = paragraph[1:]
            if not paragraph:
                break
            cut = min(width, len(paragraph))
            if len(paragraph) > width:
                breaks = [i for i in range(width // 2, min(width + 1, len(paragraph))) if paragraph[i][0].isspace()]
                if breaks:
                    cut = breaks[-1]
            chunk, paragraph = paragraph[:cut], paragraph[cut:]
            while chunk and chunk[-1][0].isspace():
                chunk.pop()
            line, color = '', None
            for char, new_color in chunk:
                if color != new_color:
                    if color:
                        line += f'</{color}>'
                    if new_color:
                        line += f'<{new_color}>'
                    color = new_color
                line += char
            if color:
                line += f'</{color}>'
            if line:
                lines.append(line)
    return lines


def analyze(data, filename):
    suffix = Path(filename).suffix.lower()
    if suffix == '.hwpx':
        paragraphs, warnings = read_hwpx(data)
    elif suffix == '.hwp':
        paragraphs, warnings = read_hwp(data)
    elif suffix == '.txt':
        paragraphs = [[['BLACK', p]] for p in data.decode('utf-8-sig').splitlines() if p.strip()]
        warnings = ['색 구분이 없는 글입니다. 아래에서 자막으로 쓸 문단을 선택해 주세요.']
    else:
        raise ValueError('HWP, HWPX, TXT 파일을 선택해 주세요.')
    entries = []
    for runs in paragraphs:
        # A color change is a caption boundary; red runs attach to adjacent marked text.
        groups = []
        color, current = None, ''
        for c, text in runs:
            if c == 'RED':
                current += '<r>' + text + '</r>'
            elif c == color or c == 'BLACK' and not text.strip():
                current += text
            else:
                if current.strip():
                    groups.append((color or 'RED', current.strip()))
                color, current = c, text
        if current.strip():
            groups.append((color or 'RED', current.strip()))
        for color, text in groups:
            text = re.sub(r'<영상\s*\d+>|[◎◯▶]', '', text).strip()
            if not text:
                continue
            kind = 'caption' if color == 'BLUE' else 'word' if color == 'GREEN' else 'skip'
            plain = re.sub(r'</?[rb]>', '', text)
            ref = REF.search(plain)
            if kind != 'skip' and ref:
                kind = 'verse'
                text = REF.sub('', text).strip()
            entries.append({'type': kind, 'text': text, 'ref': ref.group(1) if ref else '',
                            'color': color, 'speaker': '', 'reveal': False})
    if len(entries) > 3000:
        raise ValueError('문단이 3,000개를 넘습니다. 원고를 나눠 주세요.')
    return {'entries': entries, 'warnings': warnings,
            'title': Path(filename).stem, 'date': ''}


def make_spec(state):
    slides = []
    if state.get('title') or state.get('date'):
        slides.append({'type': 'cover', 'title': state.get('title', ''), 'date': state.get('date', '')})
    for entry in state['entries']:
        kind = entry['type']
        if kind == 'skip':
            continue
        if kind not in ('caption', 'word', 'verse', 'passage', 'emphasis', 'emphasis2', 'notice'):
            raise ValueError('지원하지 않는 슬라이드 종류입니다.')
        text = entry.get('text', '').strip()
        if not text:
            continue
        width = 30 if kind == 'caption' else 28 if kind == 'verse' else 25
        lines = wrap_marked(text, width)
        limit = 3 if kind == 'caption' else 6 if kind == 'verse' and entry.get('ref') else 8 if kind == 'verse' else 5
        pages = [lines] if kind == 'caption' else [lines[i:i + limit] for i in range(0, len(lines), limit)]
        for i, page in enumerate(pages):
            slide = {'type': kind, 'lines': page, 'ref': entry.get('ref', '') if i == 0 or kind == 'passage' else ''}
            if kind == 'word' and entry.get('speaker') == '성령':
                slide['speaker'] = '성령'
            if entry.get('reveal'):
                # Each manually edited line is one reveal step; stated in the UI.
                slide['chunks'] = [[line] for line in page]
            if i + 1 < len(pages):
                slide['lines'][-1] += ' <b>>></b>'
            slides.append(slide)
    if not any(s['type'] != 'cover' for s in slides):
        raise ValueError('자막으로 만들 문단이 없습니다. 문단 종류를 선택해 주세요.')
    return {'slides': slides}


def flatten(spec):
    flat = []
    for entry in spec['slides']:
        parts = template.split_caption(entry) if entry['type'] == 'caption' else [entry]
        for part in parts:
            flat.extend(template.expand_reveal(part))
    if len(flat) > 2000:
        raise ValueError('슬라이드가 2,000장을 넘습니다. 원고를 나눠 주세요.')
    return flat


def validate_template(data):
    z = checked_zip(data)
    required = ['ppt/presentation.xml', '[Content_Types].xml'] + ['ppt/slides/' + n for n in template.EXEMPLAR.values()]
    if any(n not in z.namelist() for n in required):
        raise ValueError('자막 스킬용 템플릿이 아닙니다. 원본 템플릿의 슬라이드 구조를 유지한 PPTX를 선택해 주세요.')
    return True


def build_ppt(state, template_data):
    validate_template(template_data)
    spec = make_spec(state)
    flat = flatten(spec)
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        checked_zip(template_data).extractall(root)
        styles = template.extract_styles(root)
        rel_path = root / 'ppt/_rels/presentation.xml.rels'
        rels = xml(rel_path.read_bytes())
        for e in list(rels):
            if e.get('Type', '').endswith('/slide'):
                rels.remove(e)
        pres_path = root / 'ppt/presentation.xml'
        pres = xml(pres_path.read_bytes())
        ids = pres.find(f'{{{P}}}sldIdLst')
        ids.clear()
        slide_dir = root / 'ppt/slides'
        ct_path = root / '[Content_Types].xml'
        content_types = xml(ct_path.read_bytes())
        rid_max = max([int(e.get('Id')[3:]) for e in rels if re.fullmatch(r'rId\d+', e.get('Id', ''))] + [0])
        for i, entry in enumerate(flat, 1001):
            source = slide_dir / template.EXEMPLAR[entry['type']]
            dest = slide_dir / f'slide{i}.xml'
            dest.write_bytes(source.read_bytes())
            src_rels = slide_dir / '_rels' / (source.name + '.rels')
            new_rels = xml(src_rels.read_bytes())
            for relation in list(new_rels):
                if relation.get('Type', '').endswith(('/notesSlide', '/slide')):
                    new_rels.remove(relation)
            # Keep original namespace prefixing for the supplied text/layout helpers.
            rel_string = ET.tostring(new_rels, encoding='unicode').replace('ns0:', '').replace(':ns0', '')
            (slide_dir / '_rels' / (dest.name + '.rels')).write_text(rel_string, encoding='utf-8')
            template.set_slide_text(dest, entry['type'], entry, styles)
            if entry.get('speaker'):
                template.apply_speaker_layout(root, dest, entry['speaker'])
            rid = f'rId{rid_max + i}'
            ET.SubElement(rels, f'{{{REL}}}Relationship', {'Id': rid, 'Type': R + '/slide', 'Target': f'slides/slide{i}.xml'})
            ET.SubElement(ids, f'{{{P}}}sldId', {'id': str(i), f'{{{R}}}id': rid})
            ET.SubElement(content_types, f'{{{CT}}}Override', {'PartName': f'/ppt/slides/slide{i}.xml', 'ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'})
        rel_path.write_bytes(ET.tostring(rels, encoding='utf-8', xml_declaration=True))
        pres_path.write_bytes(ET.tostring(pres, encoding='utf-8', xml_declaration=True))
        # OPC relationship graph: exclude original example slides, notes and stale previews.
        keep = {'[Content_Types].xml'}
        queue = ['']
        while queue:
            part = queue.pop()
            rp = (posixpath.dirname(part) + '/_rels/' + posixpath.basename(part) + '.rels').lstrip('/') if part else '_rels/.rels'
            if part:
                keep.add(part)
            if not (root / rp).exists() or rp in keep:
                continue
            keep.add(rp)
            relationships = xml((root / rp).read_bytes())
            for rel in list(relationships):
                if rel.get('Type', '').endswith('/thumbnail'):
                    relationships.remove(rel)
                    continue
                if rel.get('TargetMode') == 'External':
                    continue
                target = rel.get('Target', '').split('#')[0]
                resolved = posixpath.normpath(posixpath.join(posixpath.dirname(part), target)).lstrip('/')
                if resolved not in keep:
                    queue.append(resolved)
            (root / rp).write_bytes(ET.tostring(relationships, encoding='utf-8', xml_declaration=True))
        for element in list(content_types):
            if local(element.tag) == 'Override' and element.get('PartName', '').lstrip('/') not in keep:
                content_types.remove(element)
        ct_path.write_bytes(ET.tostring(content_types, encoding='utf-8', xml_declaration=True))
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for name in sorted(keep):
                if not (root / name).is_file():
                    raise ValueError('템플릿의 연결 파일을 찾지 못했습니다: ' + name)
                z.write(root / name, name)
        return out.getvalue(), len(flat)
