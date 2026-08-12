import os
import re
import zipfile
import xml.etree.ElementTree as ET

FILES = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'files')

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def path(name):
  return os.path.join(FILES, name)


def _shift_rows(text, delta):
  # Only row-relative references move when a shared formula is filled down;
  # a '$' immediately before the digits pins the row.
  return re.sub(r'(\$?[A-Z]{1,3})(\$?)(\d+)',
                lambda m: '{}{}{}'.format(
                  m.group(1), m.group(2),
                  m.group(3) if m.group(2) else int(m.group(3)) + delta),
                text)


def corrupt_formula(tmp_path, name, part='xl/worksheets/sheet1.bin',
                    find=b'\x44\x01\x00\x00\x00\x00\xc0\x44\x00\x00\x00\x00\x00\xc0\x03'):
  """Copy a workbook, replacing one formula's trailing operator with a Ptg
  that does not exist, so its token stream genuinely fails to parse."""
  src = zipfile.ZipFile(path(name))
  data = src.read(part)
  assert data.count(find) == 1, 'expected exactly one matching token stream'
  data = data.replace(find, find[:-1] + b'\x00')

  out = str(tmp_path / ('corrupt-' + name))
  with zipfile.ZipFile(out, 'w') as dst:
    for item in src.infolist():
      dst.writestr(item, data if item.filename == part else src.read(item.filename))
  return out


def xlsx_formulas(name, sheet):
  """Ground truth: the formula text Excel itself writes for the same workbook.

  Comparing against Excel's own .xlsx output keeps the expectations honest --
  hand-written ones would only ever encode our reading of the spec.
  """
  zf = zipfile.ZipFile(path(name))
  rels = {}
  for el in ET.fromstring(zf.read('xl/_rels/workbook.xml.rels')):
    rels[el.attrib['Id']] = el.attrib['Target']

  wb = ET.fromstring(zf.read('xl/workbook.xml'))
  target = None
  for el in wb.find(NS + 'sheets'):
    if el.attrib['name'] == sheet:
      target = rels[el.attrib[RNS + 'id']].lstrip('/')
      if not target.startswith('xl/'):
        target = 'xl/' + target
  if target is None:
    raise KeyError(sheet)

  out = {}
  present = set()
  shared_members = []
  hosts = {}
  array_spans = []

  root = ET.fromstring(zf.read(target))
  for row in root.find(NS + 'sheetData'):
    for c in row:
      present.add(c.attrib['r'])
      f = c.find(NS + 'f')
      if f is None:
        continue
      kind = f.attrib.get('t', 'normal')
      if f.text:
        out[c.attrib['r']] = f.text
        if kind == 'shared' and f.attrib.get('si') is not None:
          hosts[f.attrib['si']] = (c.attrib['r'], f.text)
        elif kind == 'array' and f.attrib.get('ref'):
          array_spans.append((f.attrib['ref'], f.text))
      elif kind == 'shared' and f.attrib.get('si') is not None:
        shared_members.append((c.attrib['r'], f.attrib['si']))

  # Excel writes a shared formula's text once, on the host cell; the other
  # members carry an empty <f> tagged with the same si.
  for cell, si in shared_members:
    host, text = hosts[si]
    delta = int(re.search(r'\d+$', cell).group()) - int(re.search(r'\d+$', host).group())
    out[cell] = _shift_rows(text, delta)

  # Members of a multi-cell array formula get no <f> at all, so they have to
  # come from the span -- but only for cells the sheet actually contains. A
  # span's `ref` can reach past the last populated row.
  for ref, text in array_spans:
    m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)$', ref)
    if not m:
      continue
    for r in range(int(m.group(2)), int(m.group(4)) + 1):
      cell = '{}{}'.format(m.group(1), r)
      if cell in present and cell not in out:
        out[cell] = text
  return out
