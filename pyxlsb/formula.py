"""Parsing of BIFF12 parsed-expression (Rgce) token streams into formula text.

A cell's formula is stored as a postfix stream of Ptg tokens (the ``rgce``)
plus a blob of out-of-band extras (the ``rgcb``) for the few tokens too large
to inline. This module turns that pair back into the A1-style text Excel would
show in the formula bar.

Rendering is purely structural: Excel emits an explicit PtgParen wherever the
author wrote a bracket, so re-associating the postfix stream reproduces the
original spacing of operators without needing a precedence table.

See [MS-XLSB] 2.5.97 for the token layouts.
"""

import re
import struct
import sys
from .ftab import FTAB, FTAB_ARITY

if sys.version_info > (3,):
  xrange = range

MAX_ROW = 1048576
MAX_COL = 16384

_uint16_t = struct.Struct('<H')
_uint32_t = struct.Struct('<I')
_double_t = struct.Struct('<d')

# BErr values, as used by PtgErr and by error entries inside array constants.
ERROR_CODES = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0F: '#VALUE!',
  0x17: '#REF!',
  0x1D: '#NAME?',
  0x24: '#NUM!',
  0x2A: '#N/A',
  0x2B: '#GETTING_DATA'
}

_BINARY_OPS = {
  0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&',
  0x09: '<', 0x0A: '<=', 0x0B: '=', 0x0C: '>=', 0x0D: '>', 0x0E: '<>',
  0x0F: ' ', 0x10: ',', 0x11: ':'
}

_SAFE_SHEET_NAME = re.compile(r'^[A-Za-z_][A-Za-z0-9_.]*$')
_LOOKS_LIKE_REF = re.compile(r'^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$')


class FormulaError(Exception):
  """Raised when an Rgce stream cannot be turned into formula text.

  Deliberately distinct from "this cell has no formula": a caller that treats
  a parse failure as an absent formula would mistake a computed cell for a
  hardcoded one.
  """


def col_name(idx):
  name = ''
  idx += 1
  while idx > 0:
    idx, rem = divmod(idx - 1, 26)
    name = chr(65 + rem) + name
  return name


def quote_sheet_name(name):
  if _SAFE_SHEET_NAME.match(name) and not _LOOKS_LIKE_REF.match(name):
    return name
  return "'" + name.replace("'", "''") + "'"


def _format_number(value):
  if value == int(value) and abs(value) < 1e15:
    return str(int(value))
  text = repr(float(value))
  if text.endswith('.0'):
    text = text[:-2]
  return text


def _format_string(value):
  return '"' + value.replace('"', '""') + '"'


class FormulaContext(object):
  """Workbook-level tables a formula needs in order to name what it points at.

  Without one, tokens that reference the workbook (3D refs, defined names,
  tables) have nothing to resolve against and parsing raises rather than
  inventing a placeholder.
  """

  def __init__(self, sheets=None, xtis=None, names=None, tables=None):
    super(FormulaContext, self).__init__()
    self.sheets = list(sheets or [])
    self.xtis = list(xtis or [])
    self.names = list(names or [])
    self.tables = dict(tables or {})

  def sheet_ref(self, ixti):
    """Render the `Sheet!` prefix for an external reference index."""
    if ixti < 0 or ixti >= len(self.xtis):
      return '#REF!'
    supbook, first, last = self.xtis[ixti]
    if first < 0 or last < 0:
      return '#REF!'
    prefix = '' if supbook == 0 else '[{}]'.format(supbook)
    try:
      first_name = self.sheets[first]
      last_name = self.sheets[last]
    except IndexError:
      return '#REF!'
    if first == last:
      body = first_name
    else:
      body = '{}:{}'.format(first_name, last_name)
    if prefix:
      return quote_sheet_name(prefix + body)
    return quote_sheet_name(body)

  def name(self, idx):
    if idx < 1 or idx > len(self.names):
      raise FormulaError('defined name index {} out of range'.format(idx))
    return self.names[idx - 1]

  def table(self, idx):
    if idx not in self.tables:
      raise FormulaError('table index {} not found'.format(idx))
    return self.tables[idx]


class _Buffer(object):
  """Cursor over a bytes object, in the little-endian primitives Ptgs use."""

  def __init__(self, buf):
    super(_Buffer, self).__init__()
    self._buf = buf
    self._pos = 0

  def __len__(self):
    return len(self._buf)

  @property
  def pos(self):
    return self._pos

  def eof(self):
    return self._pos >= len(self._buf)

  def read(self, size):
    if self._pos + size > len(self._buf):
      raise FormulaError('unexpected end of token stream')
    out = self._buf[self._pos:self._pos + size]
    self._pos += size
    return out

  def skip(self, size):
    self.read(size)

  def byte(self):
    return _from_byte(self.read(1))

  def short(self):
    return _uint16_t.unpack(self.read(2))[0]

  def int(self):
    return _uint32_t.unpack(self.read(4))[0]

  def double(self):
    return _double_t.unpack(self.read(8))[0]

  def short_string(self):
    cch = self.short()
    return self.read(cch * 2).decode('utf-16-le', 'replace')

  def long_string(self):
    cch = self.int()
    return self.read(cch * 2).decode('utf-16-le', 'replace')


if sys.version_info > (3,):
  def _from_byte(b):
    return b[0]
else:
  def _from_byte(b):
    return ord(b)


class Ref(object):
  """A single cell reference and its absolute/relative flags."""

  __slots__ = ('row', 'col', 'row_rel', 'col_rel')

  def __init__(self, row, col, row_rel, col_rel):
    self.row = row
    self.col = col
    self.row_rel = row_rel
    self.col_rel = col_rel

  def text(self):
    if self.row < 0 or self.col < 0 or self.row >= MAX_ROW or self.col >= MAX_COL:
      return '#REF!'
    return '{}{}{}{}'.format(
      '' if self.col_rel else '$', col_name(self.col),
      '' if self.row_rel else '$', self.row + 1)


def _read_loc(buf, base=None):
  """RgceLoc / RgceLocRel: a 4-byte row then a packed column + rel flags."""
  row = buf.int()
  packed = buf.short()
  col = packed & 0x3FFF
  col_rel = bool(packed & 0x4000)
  row_rel = bool(packed & 0x8000)
  if base is not None:
    # Shared-formula offsets: relative components are signed deltas from the
    # cell the formula was instantiated into, and wrap like Excel's do.
    if row_rel:
      if row >= MAX_ROW // 2:
        row -= MAX_ROW
      row = (base[0] + row) % MAX_ROW
    if col_rel:
      if col >= MAX_COL // 2:
        col -= MAX_COL
      col = (base[1] + col) % MAX_COL
  return Ref(row, col, row_rel, col_rel)


def _read_area(buf, base=None):
  row_first = buf.int()
  row_last = buf.int()
  first_packed = buf.short()
  last_packed = buf.short()

  col_first = first_packed & 0x3FFF
  col_last = last_packed & 0x3FFF
  first = Ref(row_first, col_first, bool(first_packed & 0x8000), bool(first_packed & 0x4000))
  last = Ref(row_last, col_last, bool(last_packed & 0x8000), bool(last_packed & 0x4000))

  if base is not None:
    for ref in (first, last):
      if ref.row_rel:
        if ref.row >= MAX_ROW // 2:
          ref.row -= MAX_ROW
        ref.row = (base[0] + ref.row) % MAX_ROW
      if ref.col_rel:
        if ref.col >= MAX_COL // 2:
          ref.col -= MAX_COL
        ref.col = (base[1] + ref.col) % MAX_COL
  return first, last


def _area_text(first, last):
  # Excel stores a whole-column reference as rows 0..MAX_ROW-1 and a whole-row
  # reference as columns 0..MAX_COL-1; both render without the spanned axis.
  if first.row == 0 and last.row == MAX_ROW - 1:
    return '{}{}:{}{}'.format(
      '' if first.col_rel else '$', col_name(first.col),
      '' if last.col_rel else '$', col_name(last.col))
  if first.col == 0 and last.col == MAX_COL - 1:
    return '{}{}:{}{}'.format(
      '' if first.row_rel else '$', first.row + 1,
      '' if last.row_rel else '$', last.row + 1)
  return '{}:{}'.format(first.text(), last.text())


class Token(object):
  """One Ptg. `render` mutates the operand stack the way Excel's would."""

  needs_extra = False

  def render(self, stack, ctx):
    raise NotImplementedError

  def read_extra(self, buf, ctx):
    pass


class Operand(Token):
  def __init__(self, text):
    super(Operand, self).__init__()
    self._text = text

  def render(self, stack, ctx):
    stack.append(self._text)


class BinaryOp(Token):
  def __init__(self, op):
    super(BinaryOp, self).__init__()
    self.op = op

  def render(self, stack, ctx):
    if len(stack) < 2:
      raise FormulaError('operand underflow at binary operator {!r}'.format(self.op))
    right = stack.pop()
    left = stack.pop()
    stack.append('{}{}{}'.format(left, self.op, right))


class UnaryOp(Token):
  def __init__(self, fmt):
    super(UnaryOp, self).__init__()
    self.fmt = fmt

  def render(self, stack, ctx):
    if not stack:
      raise FormulaError('operand underflow at unary operator')
    stack.append(self.fmt.format(stack.pop()))


class Noop(Token):
  def render(self, stack, ctx):
    pass


class RefToken(Token):
  def __init__(self, ref):
    super(RefToken, self).__init__()
    self.ref = ref

  def render(self, stack, ctx):
    stack.append(self.ref.text())


class AreaToken(Token):
  def __init__(self, first, last):
    super(AreaToken, self).__init__()
    self.first = first
    self.last = last

  def render(self, stack, ctx):
    stack.append(_area_text(self.first, self.last))


class Ref3dToken(Token):
  def __init__(self, ixti, ref, last=None, err=False):
    super(Ref3dToken, self).__init__()
    self.ixti = ixti
    self.ref = ref
    self.last = last
    self.err = err

  def render(self, stack, ctx):
    if ctx is None:
      raise FormulaError('a FormulaContext is required to resolve 3D references')
    sheet = ctx.sheet_ref(self.ixti)
    if self.err:
      body = '#REF!'
    elif self.last is not None:
      body = _area_text(self.ref, self.last)
    else:
      body = self.ref.text()
    stack.append('{}!{}'.format(sheet, body))


class NameToken(Token):
  # PtgNameX carries an ixti as well, but a name belonging to this workbook is
  # written unqualified, which is the only case we can name properly; a name
  # owned by a linked workbook would need that workbook's own name table.
  def __init__(self, idx, ixti=None):
    super(NameToken, self).__init__()
    self.idx = idx
    self.ixti = ixti

  def render(self, stack, ctx):
    if ctx is None:
      raise FormulaError('a FormulaContext is required to resolve defined names')
    stack.append(ctx.name(self.idx))


class FuncToken(Token):
  def __init__(self, iftab, argc=None):
    super(FuncToken, self).__init__()
    self.iftab = iftab
    self.argc = argc

  def render(self, stack, ctx):
    name = FTAB.get(self.iftab)
    argc = self.argc
    if argc is None:
      if self.iftab not in FTAB_ARITY:
        raise FormulaError(
          'unknown argument count for built-in function {} (iftab {})'
          .format(name or '?', self.iftab))
      argc = FTAB_ARITY[self.iftab]
    if name is None:
      raise FormulaError('unknown built-in function index {}'.format(self.iftab))
    if len(stack) < argc:
      raise FormulaError('operand underflow calling {}'.format(name))
    args = stack[len(stack) - argc:] if argc else []
    del stack[len(stack) - argc:]
    if name == 'UDF' and args:
      # iftab 255 is the user-defined-function marker: the first operand is
      # the function's own name, the rest are its arguments.
      stack.append('{}({})'.format(args[0], ','.join(args[1:])))
    else:
      stack.append('{}({})'.format(name, ','.join(args)))


class SumToken(Token):
  """PtgAttrSum: the single-argument SUM Excel special-cases."""

  def render(self, stack, ctx):
    if not stack:
      raise FormulaError('operand underflow at PtgAttrSum')
    stack.append('SUM({})'.format(stack.pop()))


class ArrayToken(Token):
  needs_extra = True

  def __init__(self):
    super(ArrayToken, self).__init__()
    self._text = None

  def read_extra(self, buf, ctx):
    rows = buf.int()
    cols = buf.int()
    grid = []
    for _ in xrange(rows):
      row = []
      for _ in xrange(cols):
        row.append(_read_ser_ar(buf))
      grid.append(row)
    self._text = '{' + ';'.join(','.join(r) for r in grid) + '}'

  def render(self, stack, ctx):
    if self._text is None:
      raise FormulaError('array constant is missing its rgcb payload')
    stack.append(self._text)


def _read_ser_ar(buf):
  kind = buf.byte()
  if kind == 0x00:
    return _format_number(buf.double())
  if kind == 0x01:
    return _format_string(buf.short_string())
  if kind == 0x02:
    return 'TRUE' if buf.byte() else 'FALSE'
  if kind == 0x03:
    return ''
  if kind == 0x04:
    return ERROR_CODES.get(buf.byte(), '#N/A')
  raise FormulaError('unknown array constant element type 0x{:02X}'.format(kind))


class MemToken(Token):
  """PtgMemArea and friends: a cached-result hint wrapping a subexpression.

  The subexpression that follows renders on its own, so the token itself
  contributes nothing beyond consuming its rgcb entry.
  """

  def __init__(self, has_extra):
    super(MemToken, self).__init__()
    self.needs_extra = has_extra

  def read_extra(self, buf, ctx):
    count = buf.int()
    buf.skip(count * 16)

  def render(self, stack, ctx):
    pass


class ExpToken(Token):
  """PtgExp: this cell defers to a shared or array formula hosted elsewhere."""

  needs_extra = True

  def __init__(self, row):
    super(ExpToken, self).__init__()
    self.row = row
    self.col = None

  def read_extra(self, buf, ctx):
    self.col = buf.int()

  def render(self, stack, ctx):
    raise FormulaError('shared formula host at row {} must be resolved by the caller'
                       .format(self.row))


class ListToken(Token):
  """PtgList: a structured reference into a table, e.g. ``MyTable[Qty]``."""

  def __init__(self, ixti, flags, list_index, col_first, col_last):
    super(ListToken, self).__init__()
    self.ixti = ixti
    self.flags = flags
    self.list_index = list_index
    self.col_first = col_first
    self.col_last = col_last

  def render(self, stack, ctx):
    if ctx is None:
      raise FormulaError('a FormulaContext is required to resolve table references')
    name, columns = ctx.table(self.list_index)
    row_type = self.flags & 0x001F
    selector = {0x02: '#All', 0x04: '#Headers', 0x08: '#Totals',
                0x10: '#This Row'}.get(row_type)

    parts = []
    if selector is not None:
      parts.append('[' + selector + ']')
    if columns and self.col_first < len(columns):
      if self.col_first == self.col_last:
        parts.append('[' + columns[self.col_first] + ']')
      elif self.col_last < len(columns):
        parts.append('[' + columns[self.col_first] + ']:[' + columns[self.col_last] + ']')
    if not parts:
      stack.append(name)
    elif len(parts) == 1 and selector is None:
      stack.append('{}{}'.format(name, parts[0]))
    else:
      stack.append('{}[{}]'.format(name, ','.join(parts)))


def parse(rgce, rgcb=b'', ctx=None, base=None):
  """Parse an Rgce/Rgcb pair into tokens.

  `base` is the (row, col) of the cell the formula lives in, required only for
  the relative-offset tokens a shared formula uses.
  """
  buf = _Buffer(rgce)
  tokens = []
  while not buf.eof():
    tokens.append(_read_token(buf, base))

  extras = _Buffer(rgcb)
  for token in tokens:
    if token.needs_extra:
      token.read_extra(extras, ctx)
  return tokens


def _read_token(buf, base):
  ptg = buf.byte()

  if ptg in _BINARY_OPS:
    return BinaryOp(_BINARY_OPS[ptg])
  if ptg == 0x12:
    return UnaryOp('+{}')
  if ptg == 0x13:
    return UnaryOp('-{}')
  if ptg == 0x14:
    return UnaryOp('{}%')
  if ptg == 0x15:
    return UnaryOp('({})')
  if ptg == 0x16:
    return Operand('')
  if ptg == 0x17:
    return Operand(_format_string(buf.short_string()))
  if ptg == 0x1C:
    return Operand(ERROR_CODES.get(buf.byte(), '#N/A'))
  if ptg == 0x1D:
    return Operand('TRUE' if buf.byte() else 'FALSE')
  if ptg == 0x1E:
    return Operand(str(buf.short()))
  if ptg == 0x1F:
    return Operand(_format_number(buf.double()))
  if ptg == 0x01:
    return ExpToken(buf.int())
  if ptg == 0x02:
    buf.skip(8)
    return Noop()
  if ptg == 0x18:
    return _read_elf(buf)
  if ptg == 0x19:
    return _read_attr(buf)

  if ptg < 0x20:
    raise FormulaError('unsupported Ptg 0x{:02X}'.format(ptg))

  # Operand Ptgs carry an operand class in the top three bits; the low five
  # select which token it is.
  base_ptg = ptg & 0x1F
  if base_ptg == 0x00:
    buf.skip(14)
    return ArrayToken()
  if base_ptg == 0x01:
    return FuncToken(buf.short())
  if base_ptg == 0x02:
    argc = buf.byte()
    return FuncToken(buf.short() & 0x7FFF, argc)
  if base_ptg == 0x03:
    return NameToken(buf.int())
  if base_ptg == 0x04:
    return RefToken(_read_loc(buf))
  if base_ptg == 0x05:
    first, last = _read_area(buf)
    return AreaToken(first, last)
  if base_ptg in (0x06, 0x07, 0x08):
    buf.skip(6)
    return MemToken(base_ptg == 0x06)
  if base_ptg == 0x09:
    buf.skip(2)
    return MemToken(False)
  if base_ptg == 0x0A:
    buf.skip(6)
    return Operand('#REF!')
  if base_ptg == 0x0B:
    buf.skip(12)
    return Operand('#REF!')
  if base_ptg == 0x0C:
    return RefToken(_read_loc(buf, base))
  if base_ptg == 0x0D:
    first, last = _read_area(buf, base)
    return AreaToken(first, last)
  if base_ptg in (0x0E, 0x0F):
    buf.skip(2)
    return MemToken(False)
  if base_ptg == 0x19:
    ixti = buf.short()
    return NameToken(buf.int(), ixti)
  if base_ptg == 0x1A:
    ixti = buf.short()
    return Ref3dToken(ixti, _read_loc(buf))
  if base_ptg == 0x1B:
    ixti = buf.short()
    first, last = _read_area(buf)
    return Ref3dToken(ixti, first, last)
  if base_ptg == 0x1C:
    ixti = buf.short()
    buf.skip(6)
    return Ref3dToken(ixti, None, err=True)
  if base_ptg == 0x1D:
    ixti = buf.short()
    buf.skip(12)
    return Ref3dToken(ixti, None, err=True)

  raise FormulaError('unsupported Ptg 0x{:02X}'.format(ptg))


def _read_attr(buf):
  flags = buf.byte()
  if flags & 0x04:
    count = buf.short()
    buf.skip((count + 1) * 2)
    return Noop()
  if flags & 0x10:
    buf.skip(2)
    return SumToken()
  buf.skip(2)
  return Noop()


# Extended (eptg) tokens introduced by ptg 0x18.
def _read_elf(buf):
  eptg = buf.byte()
  if eptg == 0x19:
    ixti = buf.short()
    flags = buf.short()
    list_index = buf.int()
    col_first = buf.short()
    col_last = buf.short()
    return ListToken(ixti, flags, list_index, col_first, col_last)
  if eptg in (0x01, 0x02, 0x03, 0x06, 0x07, 0x0A, 0x0B, 0x0D, 0x0F, 0x10, 0x1D):
    buf.skip(4)
    return Noop()
  raise FormulaError('unsupported extended Ptg 0x18 0x{:02X}'.format(eptg))


def stringify(tokens, ctx=None):
  """Render parsed tokens as formula text (without a leading ``=``)."""
  stack = []
  for token in tokens:
    token.render(stack, ctx)
  if len(stack) != 1:
    raise FormulaError('formula did not reduce to a single expression '
                       '({} operands left)'.format(len(stack)))
  return stack[0]


def to_string(rgce, rgcb=b'', ctx=None, base=None):
  return stringify(parse(rgce, rgcb, ctx, base), ctx)


def shared_host(tokens):
  """Return the (row, col) a PtgExp-only stream defers to, else None."""
  if len(tokens) == 1 and isinstance(tokens[0], ExpToken):
    return (tokens[0].row, tokens[0].col)
  return None
