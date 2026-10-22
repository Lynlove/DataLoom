"""Shared row-based chunk addressing for all sheet writers.

Cell coordinates are zero-based. Empty row bands need no physical record;
chunk_count is an exclusive index bound, not a count of stored records.
"""
from app.config import CHUNK_SIZE


def chunk_index_for_row(row: int) -> int:
    return row // CHUNK_SIZE


def chunk_row_bounds(index: int):
    return index * CHUNK_SIZE, (index + 1) * CHUNK_SIZE - 1


def empty_sheet_chunk_count(total_rows: int) -> int:
    # Preserve full-save metadata semantics even when no cells are stored.
    return chunk_index_for_row(max(total_rows, 1) - 1) + 1
