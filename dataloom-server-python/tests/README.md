# Sparse sheet chunk regression tests

From `dataloom-server-python`, install `requirements.txt` and run:

```sh
python -m unittest discover -s tests -v
CHUNK_SIZE=3 python -m unittest discover -s tests -v
```

The tests create real temporary `.xlsx` files, parse both value and formula
workbooks, and persist through SQLAlchemy into isolated in-memory SQLite.
They exercise upload → incremental edit → full replacement → deletion,
asserting exact cell contents, unique coordinates, chunk addressing and metadata.
The A1001 upload/edit test fails on the original implementation because both the
original and edited value remain in storage.

## Storage contract

- Cell coordinates remain zero-based Luckysheet `r`, `c`, `v` JSON records.
- All writers use `chunk_index = r // CHUNK_SIZE` (default size 1000).
- Each row band has inclusive bounds `[index * size, (index + 1) * size - 1]`.
  Upload's final partial band now uses the same bounds as both other writers.
- Upload/full replacement omit empty bands. Editing retains existing physical
  chunks after clearing them, as before.
- `chunk_count` is the highest stored index plus one, not the physical record
  count. With no cells, it is at least one and covers the declared row extent,
  matching the existing full-save behavior. Gaps in indices are expected.
- No API routes, request/response field names, database schema, formula value
  shape or frontend code change. The frontend's all-celldata loader reads the
  stored records in index order; it does not require contiguous indices.

These tests cover service/database behavior, not browser or HTTP upload tests.
Existing incorrectly indexed records are not migrated automatically. Reimport
an authoritative workbook to rebuild them; if duplicates already contain
conflicting values, resolve them against the authoritative source first.
