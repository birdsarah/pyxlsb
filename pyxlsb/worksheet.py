import os
import sys
import xml.etree.ElementTree as ET
from . import biff12
from . import formula
from .reader import BIFF12Reader
from collections import namedtuple

if sys.version_info > (3,):
  xrange = range

FORMULA_ERROR_MODES = ('raise', 'store')

# `f` is the cell's formula text without the leading '=', or None if the cell
# holds no formula at all. Never treat a None as "no formula" when formula
# parsing was switched off -- see Workbook(parse_formulas=...).
#
# `error` is None unless formula_errors='store' and this cell's formula could
# not be read, in which case it describes the failure. A cell that failed to
# parse still has f=None, so `f` only ever holds a real formula -- code that
# reads `f` alone can never mistake a broken formula for a hardcoded value.
Cell = namedtuple('Cell', ['r', 'c', 'v', 'f', 'error'])

class Worksheet(object):
  def __init__(self, name, fp, rels_fp=None, stringtable=None, debug=False,
               formula_context=None, formula_errors='raise'):
    super(Worksheet, self).__init__()
    if formula_errors not in FORMULA_ERROR_MODES:
      raise ValueError('formula_errors must be one of {}, got {!r}'
                       .format(', '.join(FORMULA_ERROR_MODES), formula_errors))
    self.name = name
    self._reader = BIFF12Reader(fp=fp, debug=debug)
    self._rels_fp = rels_fp
    self._rels = ET.parse(rels_fp).getroot() if rels_fp is not None else None
    self._stringtable = stringtable
    self._formula_context = formula_context
    self._formula_errors = formula_errors
    self._shared_formulas = {}
    self._array_formulas = {}
    self._formula_ranges = []
    self._data_offset = 0
    self.dimension = None
    self.cols = []
    self.rels = {}
    self.hyperlinks = {}
    self._parse()

  def __enter__(self):
    return self

  def __exit__(self, type, value, traceback):
    self.close()

  def __iter__(self):
    return self.rows()

  def _parse(self):
    if self._rels is not None:
      for el in self._rels:
        self.rels[el.attrib['Id']] = el.attrib['Target']

    for item in self._reader:
      if item[0] == biff12.DIMENSION:
        self.dimension = item[1]
      elif item[0] == biff12.COL:
        self.cols.append(item[1])
      elif item[0] == biff12.SHEETDATA:
        self._data_offset = self._reader.tell()
        if self._rels is None:
          break
      elif item[0] == biff12.HYPERLINK and self._rels is not None:
        for r in xrange(item[1].h):
          for c in xrange(item[1].w):
            self.hyperlinks[item[1].r + r, item[1].c + c] = item[1].rId

  def _host_formula(self, host):
    """Find the shared or array formula a PtgExp cell defers to.

    Excel points a PtgExp at the first cell that actually carries the formula,
    which is not always the top-left of the block it was defined over: a range
    written as AO450:AR455 can be hosted at AP450 when AO450 itself holds no
    formula. So an exact hit is only the fast path, and a miss falls back to
    asking which defined range covers the host.
    """
    fmla = self._shared_formulas.get(host)
    if fmla is not None:
      return fmla, False
    fmla = self._array_formulas.get(host)
    if fmla is not None:
      return fmla, True

    row, col = host
    for r1, c1, r2, c2, fmla, is_array in reversed(self._formula_ranges):
      if r1 <= row <= r2 and c1 <= col <= c2:
        # Memoise under the host key so each one is only searched for once.
        target = self._array_formulas if is_array else self._shared_formulas
        target[host] = fmla
        return fmla, is_array
    return None, False

  def _formula_text(self, row_num, col_num, fmla):
    """Render one cell's parsed formula, following shared formulas to their host."""
    try:
      tokens = formula.parse(fmla[0], fmla[1], self._formula_context, (row_num, col_num))
      host = formula.shared_host(tokens)
      if host is None:
        return formula.stringify(tokens, self._formula_context)

      hosted, is_array = self._host_formula(host)
      if hosted is None:
        raise formula.FormulaError(
          'no shared or array formula defined at row {} column {}'.format(host[0], host[1]))

      # An array formula shows the host's text verbatim in every cell of the
      # block; a shared formula stores its references as offsets and so
      # re-renders against whichever cell it was instantiated into.
      base = host if is_array else (row_num, col_num)
      return formula.to_string(hosted[0], hosted[1], self._formula_context, base)
    except formula.FormulaError as exc:
      raise formula.FormulaError('{}!{}{}: {}'.format(
        self.name, formula.col_name(col_num), row_num + 1, exc))

  def _resolve_formulas(self, row, row_num, pending):
    # Deferred to the end of the row because Excel writes a shared formula's
    # definition after the first cell that refers to it.
    for col_num, fmla in pending:
      try:
        row[col_num] = row[col_num]._replace(
          f=self._formula_text(row_num, col_num, fmla))
      except formula.FormulaError:
        if self._formula_errors == 'raise':
          raise
        row[col_num] = row[col_num]._replace(error=str(sys.exc_info()[1]))

  def rows(self, sparse=False):
    self._reader.seek(self._data_offset, os.SEEK_SET)
    row_num = -1
    row = None
    pending = []
    for item in self._reader:
      if item[0] == biff12.ROW and item[1].r != row_num:
        if row is not None:
          self._resolve_formulas(row, row_num, pending)
          yield row
        pending = []
        if not sparse:
          while row_num < item[1].r - 1:
            row_num += 1
            yield [Cell(row_num, i, None, None, None) for i in xrange(self.dimension.c + self.dimension.w)]
        row_num = item[1].r
        row = [Cell(row_num, i, None, None, None) for i in xrange(self.dimension.c + self.dimension.w)]
      elif item[0] == biff12.SHAREDFORMULA or item[0] == biff12.ARRAYFORMULA:
        is_array = item[0] == biff12.ARRAYFORMULA
        block = item[1]
        if is_array:
          self._array_formulas[(block.r, block.c)] = block.f
        else:
          self._shared_formulas[(block.r, block.c)] = block.f
        self._formula_ranges.append((block.r, block.c,
                                     block.r + block.h - 1, block.c + block.w - 1,
                                     block.f, is_array))
      elif item[0] >= biff12.BLANK and item[0] <= biff12.FORMULA_BOOLERR:
        if item[0] == biff12.STRING and self._stringtable is not None:
          row[item[1].c] = Cell(row_num, item[1].c, self._stringtable[item[1].v], None, None)
        else:
          row[item[1].c] = Cell(row_num, item[1].c, item[1].v, None, None)
        if item[1].f is not None and self._formula_context is not None:
          pending.append((item[1].c, item[1].f))
      elif item[0] == biff12.SHEETDATA_END:
        if row is not None:
          self._resolve_formulas(row, row_num, pending)
          yield row
        break

  def close(self):
    self._reader.close()
    if self._rels_fp is not None:
      self._rels_fp.close()
