from . import biff12
from collections import namedtuple

class Handler(object):
  def __init__(self):
    super(Handler, self).__init__()

  def read(self, reader, recid, reclen):
    if reclen > 0:
      reader.skip(reclen)


class BasicHandler(Handler):
  def __init__(self, name=None):
    super(BasicHandler, self).__init__()
    self.name = name

  def read(self, reader, recid, reclen):
    super(BasicHandler, self).read(reader, recid, reclen)
    return self.name


class StringTableHandler(Handler):
  cls = namedtuple('sst', ['count', 'uniqueCount'])

  def __init__(self):
    super(StringTableHandler, self).__init__()

  def read(self, reader, recid, reclen):
    count = reader.read_int()
    unique = reader.read_int()
    return self.cls._make([count, unique])


class StringInstanceHandler(Handler):
  cls = namedtuple('si', ['t'])

  def __init__(self):
    super(StringInstanceHandler, self).__init__()

  def read(self, reader, recid, reclen):
    reader.skip(1)
    val = reader.read_string()
    return self.cls._make([val])


class SheetHandler(Handler):
  cls = namedtuple('sheet', ['sheetId', 'rId', 'name'])

  def __init__(self):
    super(SheetHandler, self).__init__()

  def read(self, reader, recid, reclen):
    reader.skip(4)
    sheetid = reader.read_int()
    relid = reader.read_string()
    name = reader.read_string()
    return self.cls._make([sheetid, relid, name])


class DimensionHandler(Handler):
  cls = namedtuple('dimension', ['r', 'c', 'h', 'w'])

  def __init__(self):
    super(DimensionHandler, self).__init__()

  def read(self, reader, recid, reclen):
    r1 = reader.read_int()
    r2 = reader.read_int()
    c1 = reader.read_int()
    c2 = reader.read_int()
    return self.cls._make([r1, c1, r2 - r1 + 1, c2 - c1 + 1])


class ColumnHandler(Handler):
  cls = namedtuple('col', ['c1', 'c2', 'width', 'style'])

  def __init__(self):
    super(ColumnHandler, self).__init__()

  def read(self, reader, recid, reclen):
    c1 = reader.read_int()
    c2 = reader.read_int()
    width = reader.read_int() / 256
    style = reader.read_int()
    return self.cls._make([c1, c2, width, style])


class RowHandler(Handler):
  cls = namedtuple('row', ['r'])

  def __init__(self):
    super(RowHandler, self).__init__()

  def read(self, reader, recid, reclen):
    r = reader.read_int()
    return self.cls._make([r])


FORMULA_RECORDS = frozenset([
  biff12.FORMULA_STRING,
  biff12.FORMULA_FLOAT,
  biff12.FORMULA_BOOL,
  biff12.FORMULA_BOOLERR
])


class CellHandler(Handler):
  # `f` carries the raw CellParsedFormula as an (rgce, rgcb) pair, or None for
  # a cell with no formula. Turning that into text needs workbook-level tables
  # the handler cannot see, so Worksheet does the rendering.
  cls = namedtuple('c', ['c', 'v', 'f', 'style'])

  def __init__(self):
    super(CellHandler, self).__init__()

  def read(self, reader, recid, reclen):
    col = reader.read_int()
    style = reader.read_int()
    val = None
    if recid == biff12.NUM:
      val = reader.read_float()
    elif recid == biff12.BOOLERR:
      val = hex(reader.read_byte())
    elif recid == biff12.BOOL:
      val = reader.read_byte() != 0
    elif recid == biff12.FLOAT:
      val = reader.read_double()
    elif recid == biff12.STRING:
      val = reader.read_int()
    elif recid == biff12.FORMULA_STRING:
      val = reader.read_string()
    elif recid == biff12.FORMULA_FLOAT:
      val = reader.read_double()
    elif recid == biff12.FORMULA_BOOL:
      val = reader.read_byte() != 0
    elif recid == biff12.FORMULA_BOOLERR:
      val = hex(reader.read_byte())

    fmla = None
    if recid in FORMULA_RECORDS:
      reader.skip(2)
      fmla = read_parsed_formula(reader)
    return self.cls._make([col, val, fmla, style])


def read_parsed_formula(reader):
  """Read a CellParsedFormula: cce, rgce, cb, rgcb."""
  cce = reader.read_int()
  if not cce:
    return None
  rgce = reader.read(cce)
  cb = reader.read_int() or 0
  rgcb = reader.read(cb) if cb else b''
  return (rgce, rgcb)


class SharedFormulaHandler(Handler):
  cls = namedtuple('shrfmla', ['r', 'c', 'h', 'w', 'f'])

  def __init__(self):
    super(SharedFormulaHandler, self).__init__()

  def read(self, reader, recid, reclen):
    r1 = reader.read_int()
    r2 = reader.read_int()
    c1 = reader.read_int()
    c2 = reader.read_int()
    cce = reader.read_int() or 0
    rgce = reader.read(cce) if cce else b''
    return self.cls._make([r1, c1, r2 - r1 + 1, c2 - c1 + 1, (rgce, b'')])


class ArrayFormulaHandler(Handler):
  cls = namedtuple('arrfmla', ['r', 'c', 'h', 'w', 'f'])

  def __init__(self):
    super(ArrayFormulaHandler, self).__init__()

  def read(self, reader, recid, reclen):
    r1 = reader.read_int()
    r2 = reader.read_int()
    c1 = reader.read_int()
    c2 = reader.read_int()
    reader.skip(1)                 # flags
    return self.cls._make([r1, c1, r2 - r1 + 1, c2 - c1 + 1, read_parsed_formula(reader)])


class NameHandler(Handler):
  cls = namedtuple('name', ['name', 'itab', 'f'])

  def __init__(self):
    super(NameHandler, self).__init__()

  def read(self, reader, recid, reclen):
    reader.skip(4)                 # flags
    reader.skip(1)                 # chKey
    itab = reader.read_signed_int()          # -1 when scoped to the workbook
    name = reader.read_string()
    return self.cls._make([name, itab, read_parsed_formula(reader)])


class ExternSheetHandler(Handler):
  cls = namedtuple('externsheet', ['xtis'])

  def __init__(self):
    super(ExternSheetHandler, self).__init__()

  def read(self, reader, recid, reclen):
    count = reader.read_int() or 0
    xtis = []
    for _ in range(count):
      supbook = reader.read_signed_int()
      first = reader.read_signed_int()
      last = reader.read_signed_int()
      if supbook is None or first is None or last is None:
        break
      xtis.append((supbook, first, last))
    return self.cls._make([xtis])


def _valid_identifier(name):
  """Names Excel will accept for a table or column: no control characters."""
  return bool(name) and all(ch >= ' ' for ch in name)


class TableHandler(Handler):
  cls = namedtuple('table', ['id', 'name'])

  # rfx (16 bytes) then twelve fixed-size fields, then the name strings.
  ID_OFFSET = 20
  NAME_OFFSET = 64

  def __init__(self):
    super(TableHandler, self).__init__()

  def read(self, reader, recid, reclen):
    reader.seek(self.ID_OFFSET)
    tid = reader.read_int()
    reader.seek(self.NAME_OFFSET)
    name = reader.read_string()
    # Returning None drops the table from the lookup, which makes any
    # structured reference to it raise instead of rendering a wrong name.
    if tid is None or not _valid_identifier(name):
      return None
    return self.cls._make([tid, name])


class TableColumnHandler(Handler):
  cls = namedtuple('tablecolumn', ['name'])

  NAME_OFFSET = 28

  def __init__(self):
    super(TableColumnHandler, self).__init__()

  def read(self, reader, recid, reclen):
    reader.seek(self.NAME_OFFSET)
    name = reader.read_string()
    if not _valid_identifier(name):
      return None
    return self.cls._make([name])


class HyperlinkHandler(Handler):
  cls = namedtuple('hyperlink', ['r', 'c', 'h', 'w', 'rId'])

  def __init__(self):
    super(HyperlinkHandler, self).__init__()

  def read(self, reader, recid, reclen):
    r1 = reader.read_int()
    r2 = reader.read_int()
    c1 = reader.read_int()
    c2 = reader.read_int()
    rId = reader.read_string()
    return self.cls._make([r1, c1, r2 - r1 + 1, c2 - c1 + 1, rId])
