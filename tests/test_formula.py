import pytest

from conftest import corrupt_formula, path, xlsx_formulas
from pyxlsb import open_workbook
from pyxlsb.formula import (FormulaContext, FormulaError, col_name,
                            quote_sheet_name, to_string)


def read_formulas(name, sheet, **kwargs):
  out = {}
  with open_workbook(path(name), **kwargs) as wb:
    with wb.get_sheet(sheet) as ws:
      for row in ws.rows():
        for cell in row:
          if cell.f is not None:
            out['{}{}'.format(col_name(cell.c), cell.r + 1)] = cell.f
  return out


CORPUS = read_formulas('fixture.xlsb', 'Formulas')
CORPUS_TRUTH = xlsx_formulas('fixture_excel.xlsx', 'Formulas')

EXTRAS = read_formulas('fixture2.xlsb', 'Calc')
EXTRAS_TRUTH = xlsx_formulas('fixture2_excel.xlsx', 'Calc')

SPACED = read_formulas('fixture3.xlsb', 'Space')
SPACED_TRUTH = xlsx_formulas('fixture3_excel.xlsx', 'Space')

SCOPED = read_formulas('fixture4.xlsb', 'Calc')
SCOPED_TRUTH = xlsx_formulas('fixture4_excel.xlsx', 'Calc')

DELETED = read_formulas('fixture5.xlsb', 'Keep')
DELETED_TRUTH = xlsx_formulas('fixture5_excel.xlsx', 'Keep')

# Excel prefixes an unresolved user function with `_xludf.` when it writes
# .xlsx, but stores the bare name in .xlsb. We report what the file holds.
EXTRAS_TRUTH['H19'] = EXTRAS_TRUTH['H19'].replace('_xludf.', '')


@pytest.mark.parametrize('ref', sorted(CORPUS_TRUTH))
def test_formula_corpus_matches_excel(ref):
  assert CORPUS.get(ref) == CORPUS_TRUTH[ref]


@pytest.mark.parametrize('ref', sorted(EXTRAS_TRUTH))
def test_array_and_edge_cases_match_excel(ref):
  assert EXTRAS.get(ref) == EXTRAS_TRUTH[ref]


@pytest.mark.parametrize('ref', sorted(SPACED_TRUTH))
def test_author_whitespace_is_preserved(ref):
  assert SPACED.get(ref) == SPACED_TRUTH[ref]


@pytest.mark.parametrize('ref', sorted(SCOPED_TRUTH))
def test_name_scope_span_and_number_forms_match_excel(ref):
  assert SCOPED.get(ref) == SCOPED_TRUTH[ref]


@pytest.mark.parametrize('ref', sorted(DELETED_TRUTH))
def test_deleted_sheet_references_match_excel(ref):
  assert DELETED.get(ref) == DELETED_TRUTH[ref]


def test_every_formula_cell_is_found():
  assert set(CORPUS) == set(CORPUS_TRUTH)
  assert set(EXTRAS) == set(EXTRAS_TRUTH)
  assert set(SPACED) == set(SPACED_TRUTH)
  assert set(SCOPED) == set(SCOPED_TRUTH)
  assert set(DELETED) == set(DELETED_TRUTH)


class TestDefinedNameScope(object):
  """A name scoped to one sheet is written qualified when seen from another.

  The qualifier does not come from the reference's ExternSheet entry -- that
  entry carries a sentinel -- but from the scope recorded on the name itself.
  """

  def test_sheet_scoped_name_is_qualified(self):
    assert SCOPED['D1'] == 'Alpha!LocalRate'
    assert SCOPED['D2'] == 'Alpha!LocalRate*2'

  def test_workbook_scoped_name_stays_bare(self):
    assert SCOPED['D3'] == 'GlobalRate'
    assert SCOPED['D4'] == 'GlobalRate+Alpha!LocalRate'


class TestDeletedSheetReferences(object):
  """Once the sheet is gone the whole reference collapses to #REF!.

  The cell part is still recorded in the token, so rendering it anyway would
  produce `#REF!!D75` -- a string Excel never writes.
  """

  def test_reference_collapses_entirely(self):
    assert DELETED['D1'] == '#REF!'
    assert DELETED['D4'] == '#REF!+A1'

  def test_inside_a_call_or_range(self):
    assert DELETED['D2'] == 'SUM(#REF!)'
    assert DELETED['D3'] == 'IF(A1=1,#REF!,0)'

  def test_untouched_formulas_are_unaffected(self):
    assert DELETED['D5'] == 'A1+A2'


class TestSheetSpans(object):
  """A span of sheets is quoted as a whole, or not at all."""

  def test_span_needing_no_quotes(self):
    assert SCOPED['D5'] == 'SUM(Jan:Mar!A1)'
    assert SCOPED['D6'] == 'SUM(Jan:Mar!A1:A5)'

  def test_span_needing_quotes_wraps_both_ends_once(self):
    assert SCOPED['D7'] == "SUM('Q1 Data:Q4 Data'!A1)"
    assert SCOPED['D8'] == "'Q1 Data:Q4 Data'!B2"


class TestNumberFormatting(object):
  """Small magnitudes are written out in full, as Excel writes them."""

  def test_small_numbers_avoid_exponent_form(self):
    assert SCOPED['D9'] == 'A1*0.00001'
    assert SCOPED['D10'] == 'A1*0.0000001'
    assert SCOPED['D11'] == '0.000000123'

  def test_large_integers_stay_integral(self):
    assert SCOPED['D13'] == '1234567890123'
    assert SCOPED['D14'] == 'A1*1000000'


class TestWhitespace(object):
  """PtgAttrSpace records the spacing the author typed, and it round-trips."""

  def test_space_before_an_argument(self):
    assert SPACED['D1'] == 'IFERROR(A1, 0)'
    assert SPACED['D2'] == 'IFERROR(A1,  0)'

  def test_space_around_operators(self):
    assert SPACED['D5'] == 'A1 + A2'
    assert SPACED['D6'] == 'A1  +  A2'

  def test_postfix_and_prefix_operators(self):
    assert SPACED['D8'] == 'A1 %'
    assert SPACED['D9'] == '- A1'

  def test_space_inside_call_parentheses(self):
    assert SPACED['D11'] == 'SUM( A1:A5 )'
    assert SPACED['D12'] == 'SUM(A1:A5 )'

  def test_intersection_operator_is_not_decoration(self):
    # The space between two ranges is an operator, not spacing, and has to
    # survive on its own terms.
    assert SPACED['D18'] == 'SUM(A1:A3 A2:A5)'

  def test_string_literals_are_untouched(self):
    assert SPACED['D16'] == '"keep  inner  spaces"'
    assert SPACED['D17'] == 'A1&"  "&A2'


def test_simple_workbook():
  assert read_formulas('simple.xlsb', 'Sheet1') == {'B2': 'A2+A1', 'B3': 'A3+B2'}


def test_cells_without_formulas_report_none():
  with open_workbook(path('simple.xlsb')) as wb:
    with wb.get_sheet('Sheet1') as ws:
      rows = list(ws.rows())
  assert rows[0][0].v == 1.0
  assert rows[0][0].f is None
  assert rows[1][1].f == 'A2+A1'


def test_values_still_read_with_parsing_disabled():
  with open_workbook(path('simple.xlsb'), parse_formulas=False) as wb:
    assert wb.formula_context is None
    with wb.get_sheet('Sheet1') as ws:
      rows = list(ws.rows())
  assert [c.v for c in rows[1]] == [2.0, 3.0]
  assert all(c.f is None for row in rows for c in row)


def test_shared_formula_offsets_are_rebased_per_cell():
  # D95:D106 is one shared formula; each member shifts by its row offset.
  assert CORPUS['D95'] == 'A1*2+B1'
  assert CORPUS['D96'] == 'A2*2+B2'
  assert CORPUS['D106'] == 'A12*2+B12'


def test_array_formula_members_share_the_host_text():
  assert EXTRAS['F1'] == EXTRAS['F2'] == EXTRAS['F3'] == 'TRANSPOSE(A1:C1)'


def test_whole_column_and_row_references():
  assert CORPUS['D39'] == 'SUM(A:A)'
  assert CORPUS['D40'] == 'SUM(1:1)'


def test_absolute_flags_round_trip():
  assert CORPUS['D31'] == '$A$1'
  assert CORPUS['D32'] == '$A1'
  assert CORPUS['D33'] == 'A$1'
  assert CORPUS['D36'] == 'A$1:$A5'


def test_cross_sheet_references_are_qualified_and_quoted():
  assert CORPUS['D78'] == 'Data!A1'
  assert CORPUS['D82'] == "'Sheet Space'!A1"
  assert EXTRAS['H1'] == "'It''s Data'!A1"


def test_structured_table_reference():
  assert CORPUS['D118'] == 'SUM(MyTable[Qty])'
  assert CORPUS['D120'] == 'SUMPRODUCT(MyTable[Qty],MyTable[Price])'


class TestRendering(object):
  def test_col_name(self):
    assert col_name(0) == 'A'
    assert col_name(25) == 'Z'
    assert col_name(26) == 'AA'
    assert col_name(16383) == 'XFD'

  def test_quote_sheet_name(self):
    assert quote_sheet_name('Data') == 'Data'
    assert quote_sheet_name('Sheet Space') == "'Sheet Space'"
    assert quote_sheet_name("It's") == "'It''s'"
    # A name that would read as a cell reference has to be quoted too.
    assert quote_sheet_name('A1') == "'A1'"


class TestFailureModes(object):
  """A formula we cannot read must never be mistaken for an absent formula."""

  def test_unknown_ptg_raises(self):
    with pytest.raises(FormulaError):
      to_string(b'\x99\x00\x00')

  def test_truncated_stream_raises(self):
    with pytest.raises(FormulaError):
      to_string(b'\x44\x00\x00')

  def test_unbalanced_stream_raises(self):
    # Two operands and no operator to combine them.
    rgce = b'\x1e\x01\x00\x1e\x02\x00'
    with pytest.raises(FormulaError):
      to_string(rgce)

  def test_three_d_reference_without_context_raises(self):
    rgce = b'\x3a\x00\x00\x00\x00\x00\x00\x00\xc0'
    with pytest.raises(FormulaError):
      to_string(rgce)

  def test_unknown_defined_name_raises(self):
    ctx = FormulaContext(names=['OnlyOne'])
    with pytest.raises(FormulaError):
      to_string(b'\x23\x09\x00\x00\x00', ctx=ctx)

  def test_error_message_names_the_cell(self, monkeypatch):
    from pyxlsb import formula

    def boom(*args, **kwargs):
      raise FormulaError('nope')

    monkeypatch.setattr(formula, 'parse', boom)
    with open_workbook(path('simple.xlsb')) as wb:
      with wb.get_sheet('Sheet1') as ws:
        with pytest.raises(FormulaError) as excinfo:
          list(ws.rows())
    assert 'Sheet1!B2' in str(excinfo.value)

  def test_failure_is_not_silently_reported_as_no_formula(self, monkeypatch):
    from pyxlsb import formula

    monkeypatch.setattr(formula, 'stringify',
                        lambda *a, **k: (_ for _ in ()).throw(FormulaError('nope')))
    with open_workbook(path('simple.xlsb')) as wb:
      with wb.get_sheet('Sheet1') as ws:
        with pytest.raises(FormulaError):
          list(ws.rows())


def _rows_with_broken_parser(monkeypatch, **kwargs):
  from pyxlsb import formula

  def boom(*args, **kw):
    raise FormulaError('synthetic failure')

  monkeypatch.setattr(formula, 'stringify', boom)
  monkeypatch.setattr(formula, 'to_string', boom)
  with open_workbook(path('simple.xlsb'), **kwargs) as wb:
    with wb.get_sheet('Sheet1') as ws:
      return list(ws.rows())


class TestStoredErrors(object):
  """formula_errors='store' keeps a bad cell from aborting the whole sheet."""

  def test_rows_still_stream_and_values_survive(self, monkeypatch):
    rows = _rows_with_broken_parser(monkeypatch, formula_errors='store')
    assert [c.v for c in rows[1]] == [2.0, 3.0]

  def test_error_is_recorded_on_the_offending_cell(self, monkeypatch):
    rows = _rows_with_broken_parser(monkeypatch, formula_errors='store')
    broken = rows[1][1]
    assert broken.error is not None
    assert 'synthetic failure' in broken.error
    assert 'Sheet1!B2' in broken.error

  def test_a_failed_cell_never_looks_like_a_hardcoded_value(self, monkeypatch):
    # The whole point of the separate field: `f` stays None *and* `error` is
    # set, so "f is None and error is None" still means "no formula here".
    rows = _rows_with_broken_parser(monkeypatch, formula_errors='store')
    assert rows[1][1].f is None
    assert rows[0][0].f is None and rows[0][0].error is None

  def test_cells_without_formulas_have_no_error(self, monkeypatch):
    rows = _rows_with_broken_parser(monkeypatch, formula_errors='store')
    assert all(c.error is None for c in rows[0])

  def test_default_mode_still_raises(self, monkeypatch):
    with pytest.raises(FormulaError):
      _rows_with_broken_parser(monkeypatch)

  def test_healthy_workbook_records_no_errors(self):
    with open_workbook(path('fixture.xlsb'), formula_errors='store') as wb:
      with wb.get_sheet('Formulas') as ws:
        errors = [c.error for row in ws.rows() for c in row if c.error]
    assert errors == []

  @pytest.mark.parametrize('mode', ['ignore', 'coerce', '', None, True])
  def test_unknown_mode_is_rejected(self, mode):
    with pytest.raises(ValueError):
      open_workbook(path('simple.xlsb'), formula_errors=mode)


class TestSharedFormulaHostLookup(object):
  """A PtgExp host is not always the top-left of the block it belongs to.

  Seen in the wild: a shared formula defined over AO450:AR455 whose cells all
  point at AP450, because AO450 carries no formula of its own. Keying the
  lookup on the range's top-left alone silently loses every cell of the block.
  """

  SHARED = (b'shared-rgce', b'')
  ARRAY = (b'array-rgce', b'')

  @pytest.fixture
  def sheet(self):
    with open_workbook(path('simple.xlsb')) as wb:
      with wb.get_sheet('Sheet1') as ws:
        # AO450:AR455, hosted at AP450 rather than its top-left AO450.
        ws._shared_formulas[(449, 40)] = self.SHARED
        ws._formula_ranges.append((449, 40, 454, 43, self.SHARED, False))
        ws._array_formulas[(10, 5)] = self.ARRAY
        ws._formula_ranges.append((10, 5, 12, 5, self.ARRAY, True))
        yield ws

  def test_exact_top_left_still_hits(self, sheet):
    assert sheet._host_formula((449, 40)) == (self.SHARED, False)

  def test_host_inside_the_range_is_found(self, sheet):
    assert sheet._host_formula((449, 41)) == (self.SHARED, False)
    assert sheet._host_formula((454, 43)) == (self.SHARED, False)

  def test_host_outside_every_range_is_not_found(self, sheet):
    assert sheet._host_formula((449, 44)) == (None, False)
    assert sheet._host_formula((455, 40)) == (None, False)
    assert sheet._host_formula((448, 40)) == (None, False)

  def test_array_blocks_are_reported_as_arrays(self, sheet):
    assert sheet._host_formula((11, 5)) == (self.ARRAY, True)

  def test_containment_hit_is_memoised(self, sheet):
    assert (449, 41) not in sheet._shared_formulas
    sheet._host_formula((449, 41))
    # Cached under the host key, so repeat lookups skip the range scan.
    assert sheet._shared_formulas[(449, 41)] == self.SHARED


class TestGenuinelyUnreadableFormula(object):
  """The same behaviour, driven by a real file rather than a patched parser.

  simple.xlsb's B2 is `A2+A1`, ending in a PtgAdd (0x03). Rewriting that byte
  to an unallocated Ptg makes the token stream genuinely unreadable, so the
  failure comes out of the parser itself.
  """

  @pytest.fixture
  def corrupted(self, tmp_path):
    return corrupt_formula(tmp_path, 'simple.xlsb')

  def test_raises_by_default(self, corrupted):
    with pytest.raises(FormulaError) as excinfo:
      with open_workbook(corrupted) as wb:
        with wb.get_sheet('Sheet1') as ws:
          list(ws.rows())
    assert 'Sheet1!B2' in str(excinfo.value)

  def test_store_mode_records_and_carries_on(self, corrupted):
    with open_workbook(corrupted, formula_errors='store') as wb:
      with wb.get_sheet('Sheet1') as ws:
        rows = list(ws.rows())

    broken = rows[1][1]
    assert broken.f is None
    assert broken.error is not None and 'Sheet1!B2' in broken.error
    assert broken.v == 3.0                      # the cached value survives

    # The undamaged formula two rows down still parses.
    assert rows[2][1].f == 'A3+B2'
    assert rows[2][1].error is None
