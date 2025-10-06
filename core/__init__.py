
from .backend import (
    guess_mapping, prepare_dataframe, auto_detect_group_fields,
    parse_date, format_date_dmy, slugify, find_target_table, clear_table_keep_header,
    fill_table, list_candidate_tables, read_placeholders_from_excel
)
from .funcionalidades import (
    build_letter_bytes, generate_letters_per_group, make_zip, build_index_sheet
)

__all__ = [
    "guess_mapping", "prepare_dataframe", "auto_detect_group_fields",
    "parse_date", "format_date_dmy", "slugify", "find_target_table", "clear_table_keep_header",
    "fill_table", "list_candidate_tables", "read_placeholders_from_excel",
    "build_letter_bytes", "generate_letters_per_group", "make_zip", "build_index_sheet",
]
