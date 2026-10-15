import io
import posixpath
import struct
import unittest
import zipfile
import zlib
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree as ET
import engine

ROOT = Path(__file__).resolve().parents[1]


def hwpx():
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w') as z:
        z.writestr('Contents/header.xml', '<head><charPr id="0" textColor="#000000"/><charPr id="1" textColor="#0000FF"/><charPr id="2" textColor="#008000"/><charPr id="3" textColor="#FF0000"/></head>')
        z.writestr('Contents/section0.xml', '<sec><p><run charPrIDRef="0"><t>이 문단은 설명입니다.</t></run></p><p><run charPrIDRef="1"><t>오늘은 </t></run><run charPrIDRef="3"><t>감사</t></run><run charPrIDRef="1"><t>하며 살자.</t></run></p><p><run charPrIDRef="2"><t>“진리의 성령이 오시면” (요 16:13)</t></run></p><p><run charPrIDRef="2"><t>“말씀을 귀하게 여겨라.”</t></run></p></sec>')
    return out.getvalue()


def record(tag, data):
    return struct.pack('<I', tag | len(data) << 20) + data


class EngineTests(unittest.TestCase):
    def test_hwpx_color_red_and_reference(self):
        result = engine.analyze(hwpx(), '원고.hwpx')
        entries = result['entries']
        self.assertEqual([e['type'] for e in entries], ['skip', 'caption', 'verse', 'word'])
        self.assertEqual(entries[1]['text'], '오늘은 <r>감사</r>하며 살자.')
        self.assertEqual(entries[2]['ref'], '요 16:13')
        self.assertNotIn('16:13', entries[2]['text'])

    def test_hwp_records_compression_and_character_offsets(self):
        header = bytearray(256)
        header[:17] = b'HWP Document File'
        struct.pack_into('<I', header, 36, 1)
        shapes = b''.join(record(21, b'\0' * 52 + bytes(rgb) + b'\0' * 13)
                          for rgb in [(0,0,0),(0,0,255),(255,0,0)])
        # Tab is eight UTF-16 units; following character offsets must retain that space.
        text = '가'.encode('utf-16le') + struct.pack('<8H', 9,0,0,0,0,0,0,9) + '나다\r'.encode('utf-16le')
        body = record(66, b'\0'*22) + record(67, text) + record(68, struct.pack('<6I', 0,1,9,2,10,1))
        def compressed(data):
            encoder = zlib.compressobj(wbits=-15)
            return encoder.compress(data) + encoder.flush()
        streams = {'FileHeader': bytes(header), 'DocInfo': compressed(shapes), 'BodyText/Section0': compressed(body)}
        class FakeOle:
            def __init__(self, *_): pass
            def __enter__(self): return self
            def __exit__(self, *_): pass
            def openstream(self, name): return io.BytesIO(streams['/'.join(name) if isinstance(name, list) else name])
            def listdir(self): return [['BodyText','Section0']]
        with patch('olefile.OleFileIO', FakeOle):
            result = engine.analyze(b'fixture', '원고.hwp')
        self.assertEqual(result['entries'][0]['text'], '가 <r>나</r>다')

    def test_plain_text_is_not_silently_selected(self):
        state = engine.analyze('검정 글씨만 있습니다.'.encode(), '원고.txt')
        self.assertEqual(state['entries'][0]['type'], 'skip')
        with self.assertRaises(ValueError): engine.make_spec(state)

    def test_xml_entities_rejected(self):
        with self.assertRaises(ValueError): engine.xml(b'<!DOCTYPE x [<!ENTITY e "bad">]><x>&e;</x>')

    def test_bad_zip_path_rejected(self):
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w') as z: z.writestr('../oops', 'x')
        with self.assertRaises(ValueError): engine.checked_zip(out.getvalue())

    def test_red_survives_wrapping(self):
        lines = engine.wrap_marked('시작 <r>' + '가' * 55 + '</r> 끝', 25)
        for line in lines:
            self.assertEqual(line.count('<r>'), line.count('</r>'))
        self.assertEqual(''.join(__import__('re').sub(r'</?r>', '', line) for line in lines).replace(' ', ''), ('시작'+'가'*55+'끝'))

    def test_template_and_all_types_produce_valid_relationships(self):
        state = engine.analyze(hwpx(), '원고.hwpx')
        state['title'] = '테스트 제목'; state['date'] = '2026년 10월 18일 주일말씀'
        state['entries'][1]['text'] = '첫째 줄\n둘째 줄\n셋째 줄\n넷째 줄'
        state['entries'][1]['reveal'] = True
        state['entries'][3]['speaker'] = '성령'
        for kind in ('passage', 'emphasis', 'emphasis2', 'notice'):
            state['entries'].append({'type':kind,'text':'검증용 <r>강조</r> 문장', 'ref':'요 16:13'})
        data, count = engine.build_ppt(state, (ROOT/'assets/template.pptx').read_bytes())
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            names = set(z.namelist())
            self.assertEqual(count, 11)
            pres = ET.fromstring(z.read('ppt/presentation.xml'))
            self.assertEqual(len(pres.find(f'{{{engine.P}}}sldIdLst')), count)
            self.assertNotIn('ppt/slides/slide1.xml', names)
            self.assertFalse(any('notesSlides' in name for name in names))
            all_text = ''.join(z.read(n).decode() for n in names if n.startswith('ppt/slides/') and n.endswith('.xml'))
            self.assertIn('테스트 제목', all_text)
            self.assertIn('FF0000', all_text)
            self.assertNotIn('이 문단은 설명입니다', all_text)
            self.assertNotIn('&lt;r&gt;', all_text)
            for name in names:
                if name.endswith(('.xml', '.rels')): ET.fromstring(z.read(name))
                if name.endswith('.rels'):
                    parent = '' if name == '_rels/.rels' else posixpath.dirname(posixpath.dirname(name))
                    for rel in ET.fromstring(z.read(name)):
                        if rel.get('TargetMode') == 'External': continue
                        target = posixpath.normpath(posixpath.join(parent, rel.get('Target'))).lstrip('/')
                        self.assertIn(target, names, f'{name} → {target}')
            layouts = [z.read(n).decode() for n in names if n.startswith('ppt/slides/_rels')]
            self.assertTrue(any('slideLayout83.xml' in r for r in layouts))


if __name__ == '__main__': unittest.main()
