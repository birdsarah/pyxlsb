import os
import sys
import xml.etree.ElementTree as ET
from . import biff12
from .formula import FormulaContext
from .reader import BIFF12Reader
from .stringtable import StringTable
from .worksheet import FORMULA_ERROR_MODES, Worksheet
from tempfile import TemporaryFile

if sys.version_info > (3,):
  basestring = (str, bytes)

class Workbook(object):
  def __init__(self, fp, debug=False, parse_formulas=True, formula_errors='raise'):
    super(Workbook, self).__init__()
    if formula_errors not in FORMULA_ERROR_MODES:
      raise ValueError('formula_errors must be one of {}, got {!r}'
                       .format(', '.join(FORMULA_ERROR_MODES), formula_errors))
    self._zf = fp
    self._debug = debug
    self._parse_formulas = parse_formulas
    self._formula_errors = formula_errors
    self._sheets = []
    self.stringtable = None
    self.formula_context = None
    self._parse()
    if parse_formulas:
      self.formula_context = self._build_formula_context()

  def __enter__(self):
    return self

  def __exit__(self, type, value, traceback):
    self.close()

  @property
  def sheets(self):
    return [v[0] for v in self._sheets]

  def _parse(self):
    rels = {}
    with self._zf.open('xl/_rels/workbook.bin.rels', 'r') as zf:
      for el in ET.parse(zf).getroot():
        rels[el.attrib['Id']] = el.attrib['Target']

    with TemporaryFile() as temp:
      with self._zf.open('xl/workbook.bin', 'r') as zf:
        temp.write(zf.read())
        temp.seek(0, os.SEEK_SET)
      reader = BIFF12Reader(fp=temp, debug=self._debug)
      for item in reader:
        if item[0] == biff12.SHEET:
          self._sheets.append((item[1].name, rels[item[1].rId]))
        elif item[0] == biff12.SHEETS_END:
          break

    try:
      temp = TemporaryFile()
      with self._zf.open('xl/sharedStrings.bin', 'r') as zf:
        temp.write(zf.read())
        temp.seek(0, os.SEEK_SET)
      self.stringtable = StringTable(fp=temp)
    except KeyError:
      temp.close()
    except Exception:
      temp.close()
      raise

  def _copy_part(self, name):
    temp = TemporaryFile()
    with self._zf.open(name, 'r') as zf:
      temp.write(zf.read())
      temp.seek(0, os.SEEK_SET)
    return temp

  def _build_formula_context(self):
    """Collect the workbook-level tables formula text is written in terms of.

    3D references arrive as an index into the ExternSheet table, defined names
    as an index into the name table, and structured references as a table id,
    so none of them can be rendered from the worksheet part alone.
    """
    xtis = []
    names = []
    name_scopes = []
    with self._copy_part('xl/workbook.bin') as temp:
      reader = BIFF12Reader(fp=temp, debug=self._debug)
      for item in reader:
        if item[0] == biff12.EXTERNSHEET:
          xtis.extend(item[1].xtis)
        elif item[0] == biff12.DEFINEDNAME:
          names.append(item[1].name)
          name_scopes.append(item[1].itab)
        elif item[0] == biff12.WORKBOOK_END:
          break

    tables = {}
    for part in self._zf.namelist():
      if not part.startswith('xl/tables/') or not part.endswith('.bin'):
        continue
      table_id = None
      table_name = None
      columns = []
      with self._copy_part(part) as temp:
        reader = BIFF12Reader(fp=temp, debug=self._debug)
        for item in reader:
          if item[0] == biff12.TABLE:
            table_id = item[1].id
            table_name = item[1].name
          elif item[0] == biff12.TABLECOLUMN:
            columns.append(item[1].name)
      if table_id is not None and table_name:
        tables[table_id] = (table_name, columns)

    return FormulaContext(sheets=self.sheets, xtis=xtis, names=names,
                          name_scopes=name_scopes, tables=tables)

  def get_sheet(self, idx, rels=False):
    if isinstance(idx, basestring):
      idx = [s.lower() for s, _ in self._sheets].index(idx.lower()) + 1
    if idx < 1 or idx > len(self._sheets):
      raise IndexError('sheet index out of range')

    name = self._sheets[idx - 1][0]
    target = self._sheets[idx - 1][1].split('/')

    temp = TemporaryFile()
    with self._zf.open('xl/{}/{}'.format(target[0], target[-1]), 'r') as zf:
      temp.write(zf.read())
      temp.seek(0, os.SEEK_SET)

    if rels:
      rels_temp = TemporaryFile()
      with self._zf.open('xl/{}/_rels/{}.rels'.format(target[0], target[-1]), 'r') as zf:
        rels_temp.write(zf.read())
        rels_temp.seek(0, os.SEEK_SET)
    else:
      rels_temp = None

    return Worksheet(name=name, fp=temp, rels_fp=rels_temp, stringtable=self.stringtable,
                     debug=self._debug, formula_context=self.formula_context,
                     formula_errors=self._formula_errors)

  def close(self):
    self._zf.close()
    if self.stringtable is not None:
      self.stringtable.close()
