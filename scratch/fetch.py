"""Temporary: fetch Council reports, decision letters, and the files they link to, and save page text.

Reads scratch/targets.json. For each item, downloads the Internet Archive copy when there is one,
otherwise RIDOH's Drive file. Saves text per page (OCR when a page has no text), the links inside each
PDF with their page number, and then follows Drive and Google Docs links one level down.
Videos and anything over 60 MB are skipped. Output goes to scratch/out/.
"""
import json, os, re, subprocess, sys, time, html
import requests
from pypdf import PdfReader

OUT = 'scratch/out'
os.makedirs(OUT, exist_ok=True)
S = requests.Session()
S.headers['User-Agent'] = 'Mozilla/5.0 (abc-ri promise sources)'
MAX = 60_000_000


def get(url):
    r = S.get(url, stream=True, timeout=120, allow_redirects=True)
    head = next(r.iter_content(4096), b'')
    ctype = r.headers.get('content-type', '')
    size = int(r.headers.get('content-length') or 0)
    if size > MAX or b'ftyp' in head[:12] or ctype.startswith('video') or ctype.startswith('audio'):
        r.close()
        return None, f'skipped (video or too big: {ctype} {size})'
    data = head + b''.join(r.iter_content(1 << 16))
    if len(data) > MAX:
        return None, 'skipped (too big)'
    return (data, ctype), None


def drive(fid):
    """Download a public Drive file, getting past the virus-scan page for big files."""
    url = f'https://drive.usercontent.google.com/download?id={fid}&export=download&confirm=t'
    res, err = get(url)
    if err:
        return None, err
    data, ctype = res
    if data[:5] == b'%PDF-' or 'html' not in ctype:
        return res, None
    page = data.decode('utf8', 'ignore')
    form = re.search(r'<form[^>]+action="([^"]+)"(.*?)</form>', page, re.S)
    if form:
        fields = dict(re.findall(r'name="([^"]+)" value="([^"]*)"', form.group(2)))
        res, err = get(html.unescape(form.group(1)) + '?' + '&'.join(f'{k}={v}' for k, v in fields.items()))
        if err:
            return None, err
        if res[0][:5] == b'%PDF-' or 'html' not in res[1]:
            return res, None
    return None, 'Drive returned a web page, not the file (private or sign-in needed)'


def gdoc(kind, did):
    fmt = {'presentation': 'pdf', 'document': 'pdf', 'spreadsheets': 'pdf'}[kind]
    return get(f'https://docs.google.com/{kind}/d/{did}/export?format={fmt}' if kind != 'presentation'
               else f'https://docs.google.com/presentation/d/{did}/export/pdf')


def pdf_text(path):
    m = re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo', path], capture_output=True, text=True).stdout)
    if not m:
        raise ValueError('PDF would not open')
    n = int(m.group(1))
    pages = []
    for p in range(1, n + 1):
        t = subprocess.run(['pdftotext', '-layout', '-f', str(p), '-l', str(p), path, '-'], capture_output=True, text=True).stdout
        if len(t.strip()) < 40 and n <= 60:
            subprocess.run(['pdftoppm', '-r', '200', '-f', str(p), '-l', str(p), '-png', path, '/tmp/pg'], check=False)
            imgs = sorted(f for f in os.listdir('/tmp') if f.startswith('pg') and f.endswith('.png'))
            if imgs:
                t = '[OCR]\n' + subprocess.run(['tesseract', '/tmp/' + imgs[-1], '-', '--psm', '3'], capture_output=True, text=True).stdout
                for f in imgs:
                    os.remove('/tmp/' + f)
        pages.append(t)
    return pages


def pdf_links(path):
    out = []
    try:
        for i, page in enumerate(PdfReader(path).pages, 1):
            for a in page.get('/Annots') or []:
                a = a.get_object()
                uri = (a.get('/A') or {}).get('/URI')
                if uri:
                    out.append({'page': i, 'uri': str(uri)})
    except Exception as e:
        out.append({'error': str(e)})
    return out


def save(key, meta, data):
    base = os.path.join(OUT, re.sub(r'[^\w-]', '_', key)[:80])
    if data[:5] == b'%PDF-':
        open(base + '.pdf', 'wb').write(data)
        try:
            pages = pdf_text(base + '.pdf')
        except Exception as e:
            meta['note'] = f'could not read PDF: {e}'
            json.dump(meta, open(base + '.json', 'w'), indent=1)
            os.remove(base + '.pdf')
            return meta
        meta['pages'] = len(pages)
        meta['links'] = pdf_links(base + '.pdf')
        with open(base + '.txt', 'w') as f:
            for i, t in enumerate(pages, 1):
                f.write(f'\n===== PAGE {i} =====\n{t}')
        os.remove(base + '.pdf')
    else:
        meta['note'] = f'not a PDF ({data[:20]!r})'
    json.dump(meta, open(base + '.json', 'w'), indent=1)
    return meta


def already(key):
    p = os.path.join(OUT, re.sub(r'[^\w-]', '_', key)[:80] + '.json')
    return json.load(open(p)) if os.path.exists(p) else None


def fetch(item):
    prev = already(item['id'])
    if prev and (prev.get('status') == 'ok' or 'skipped' in str(prev.get('status'))):
        return prev
    res, err = (None, 'no source')
    if item.get('archive_url'):
        res, err = get(item['archive_url'])
        if not err and res[0][:5] != b'%PDF-' and 'html' in res[1]:
            res, err = None, 'archive copy not a file'
    if (res is None) and item.get('kind') == 'file' and re.match(r'^[\w-]{20,}$', item['id']):
        res, err = drive(item['id'])
    if res is None:
        item['status'] = err
        json.dump(item, open(os.path.join(OUT, re.sub(r'[^\w-]', '_', item['id'])[:80] + '.json'), 'w'), indent=1)
        return item
    item['status'] = 'ok'
    return save(item['id'], item, res[0])


targets = json.load(open('scratch/targets.json'))
seen = set()
done = []
for t in targets:
    if t['id'] in seen or t['kind'] == 'folder':
        continue
    seen.add(t['id'])
    print('fetch', t['class'], t['label'][:70], flush=True)
    done.append(fetch(t))
    time.sleep(1)

# one level down: links inside Council reports
for d in list(done):
    if d.get('class') not in ('council_report', 'other'):
        continue
    for ln in d.get('links', []):
        u = ln.get('uri', '')
        m = re.search(r'drive\.google\.com/(?:file/d/|open\?id=)([\w-]{20,})', u)
        g = re.search(r'docs\.google\.com/(presentation|document|spreadsheets)/d/([\w-]{20,})', u)
        fid = m.group(1) if m else (g.group(2) if g else None)
        if not fid or fid in seen:
            continue
        seen.add(fid)
        item = {'id': fid, 'kind': 'file', 'class': 'linked', 'label': f'linked from {d["label"][:80]} (page {ln["page"]})',
                'rows': d.get('rows', []), 'from': d['id'], 'uri': u}
        print('follow', item['label'][:90], flush=True)
        if already(fid):
            continue
        if g:
            res, err = gdoc(g.group(1), fid)
            if res is None:
                item['status'] = err
                json.dump(item, open(os.path.join(OUT, fid + '.json'), 'w'), indent=1)
            else:
                item['status'] = 'ok'
                save(fid, item, res[0])
        else:
            fetch(item)
        time.sleep(1)
    if any('folders/' in ln.get('uri', '') for ln in d.get('links', [])):
        d.setdefault('folders', [ln['uri'] for ln in d['links'] if 'folders/' in ln.get('uri', '')])
        json.dump(d, open(os.path.join(OUT, re.sub(r'[^\w-]', '_', d['id'])[:80] + '.json'), 'w'), indent=1)
print('done')


# Presentations: PowerPoint files from Drive, and Google Slides exported as plain text.
import io, zipfile
def pptx_text(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    slides = sorted((n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)), key=lambda n: int(re.search(r'(\d+)', n.split('/')[-1]).group(1)))
    out = []
    for i, n in enumerate(slides, 1):
        x = z.read(n).decode('utf8', 'ignore')
        out.append(f'\n===== SLIDE {i} =====\n' + '\n'.join(html.unescape(t) for t in re.findall(r'<a:t>([^<]*)</a:t>', x)))
    return ''.join(out), len(slides)

for f in sorted(os.listdir(OUT)):
    if not f.endswith('.json'):
        continue
    m = json.load(open(os.path.join(OUT, f)))
    note = str(m.get('note', ''))
    if m.get('status') != 'ok' or not (note.startswith('not a PDF') or note.startswith('could not read PDF')) or m.get('slides'):
        continue
    u = m.get('uri', '')
    g = re.search(r'docs\.google\.com/presentation/d/([\w-]{20,})', u)
    try:
        if g:
            r = S.get(f'https://docs.google.com/presentation/d/{g.group(1)}/export/pptx', timeout=120)
            data = r.content
        else:
            res, err = drive(m['id'])
            data = res[0] if res else b''
        if data[:2] != b'PK':
            m['note'] = note + ' | retry: not a PowerPoint file'
        else:
            text, n = pptx_text(data)
            open(os.path.join(OUT, f[:-5] + '.txt'), 'w').write(text)
            m['slides'] = n
            m['note'] = 'PowerPoint text'
    except Exception as e:
        m['note'] = note + f' | retry failed: {e}'
    print('slides', m['id'], m.get('slides'), m['note'][:60], flush=True)
    json.dump(m, open(os.path.join(OUT, f), 'w'), indent=1)
    time.sleep(1)
