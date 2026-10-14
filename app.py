from calendar import monthrange
from hashlib import sha256
from io import BytesIO, StringIO
import math
import re
from typing import Any

import pandas as pd
import streamlit as st


SOURCE_COLUMN = "Source workbook"
MODIFIED_INDEX_COLUMN = "__modified_date_index__"
EXTENSION_INDEX_COLUMN = "__normalized_extension__"
SIZE_MB_COLUMN = "__size_mb__"
STATUS_COLUMN = "Activity status"
SIZE_CATEGORY_COLUMN = "Data category"
EXTENSION_CATEGORY_COLUMN = "Extension category"

ACTIVITY_CATEGORIES = ["Active", "Inactive", "Missing modified date"]
SIZE_CATEGORIES = ["Contains data", "0 file size", "Size unknown"]
EXTENSION_CATEGORIES = [
    "Business files",
    "Media files",
    "ROT files",
    "Undetermined files",
]

DEFAULT_EXTENSIONS = {
    "Business files": [
        ".accdb", ".csv", ".doc", ".docm", ".docx", ".dotx", ".dwg", ".dxf",
        ".eml", ".html", ".md", ".mdb", ".msg", ".odf", ".odp", ".ods", ".odt",
        ".one", ".pdf", ".ppt", ".pptm", ".pptx", ".pub", ".rtf", ".sql", ".txt",
        ".vsd", ".vsdx", ".xls", ".xlsb", ".xlsm", ".xlsx", ".xltx", ".xml",
    ],
    "Media files": [
        ".3gp", ".aac", ".aiff", ".avi", ".bmp", ".flac", ".gif", ".heic",
        ".jpeg", ".jpg", ".m4a", ".m4v", ".mkv", ".mov", ".mp3", ".mp4",
        ".mpeg", ".mpg", ".png", ".psd", ".raw", ".svg", ".tif", ".tiff",
        ".wav", ".webm", ".webp", ".wma", ".wmv",
    ],
    "ROT files": [
        ".bak", ".cache", ".cat", ".chk", ".class", ".com", ".dat", ".dll",
        ".drv", ".etl", ".evtx", ".exe", ".ilk", ".jar", ".log", ".lock",
        ".lnk", ".manifest", ".msi", ".msp", ".old", ".pdb", ".reg", ".sys",
        ".tmp", ".temp",
    ],
}

NIST_NSRL_URL = (
    "https://www.nist.gov/itl/csd/secure-systems-and-applications/"
    "national-software-reference-library-nsrl"
)


def normalized_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).strip().rstrip("|").casefold())


def normalize_extension(value: Any) -> str:
    if pd.isna(value):
        return ""
    extension = str(value).strip().casefold()
    if not extension or extension in {"nan", "none", "n/a"}:
        return ""
    if not extension.startswith("."):
        extension = f".{extension}"
    return extension


def add_months(value: pd.Timestamp, months: int) -> pd.Timestamp:
    month_index = value.year * 12 + value.month - 1 + months
    year, month_offset = divmod(month_index, 12)
    month = month_offset + 1
    day = min(value.day, monthrange(year, month)[1])
    return pd.Timestamp(year, month, day, tz=value.tz)


def find_column(columns: list[Any], aliases: set[str]) -> Any | None:
    for column in columns:
        if normalized_header(column) in aliases:
            return column
    return None


def inspect_workbooks(
    files: tuple[tuple[str, bytes], ...],
) -> tuple[tuple[str, ...], ...]:
    workbook_sheets = []
    for filename, contents in files:
        engine = "xlrd" if filename.lower().endswith(".xls") else "openpyxl"
        try:
            with pd.ExcelFile(BytesIO(contents), engine=engine) as workbook:
                sheet_names = tuple(workbook.sheet_names)
        except (ImportError, OSError, ValueError) as error:
            raise ValueError(f"Could not open '{filename}': {error}") from error
        if not sheet_names:
            raise ValueError(f"'{filename}' does not contain any worksheets.")
        workbook_sheets.append(sheet_names)
    return tuple(workbook_sheets)


def load_and_index_files(
    files: tuple[tuple[str, bytes, str], ...],
) -> dict[str, Any]:
    frames: list[pd.DataFrame] = []
    source_columns: list[Any] | None = None
    modified_column = None
    extension_column = None
    filename_column = None
    size_mb_column = None
    size_bytes_column = None

    for filename, contents, sheet_name in files:
        engine = "xlrd" if filename.lower().endswith(".xls") else "openpyxl"
        try:
            with pd.ExcelFile(BytesIO(contents), engine=engine) as workbook:
                frame = pd.read_excel(workbook, sheet_name=sheet_name)
        except (ImportError, OSError, ValueError, pd.errors.ParserError) as error:
            raise ValueError(
                f"Could not read worksheet '{sheet_name}' from '{filename}': {error}"
            ) from error

        if not frame.columns.is_unique:
            raise ValueError(f"'{filename}' contains duplicate column names.")
        columns = list(frame.columns)
        if source_columns is None:
            source_columns = columns
            modified_column = find_column(columns, {"modified", "modifieddate"})
            extension_column = find_column(
                columns, {"fileextension", "extension", "ext"}
            )
            filename_column = find_column(columns, {"filename", "name"})
            size_mb_column = find_column(
                columns, {"filesizemb", "sizeinmb", "sizemb"}
            )
            size_bytes_column = find_column(
                columns, {"filesize", "filesizebytes", "sizebytes"}
            )
            if modified_column is None:
                raise ValueError(
                    f"'{filename}' needs a Modified column "
                    "(a trailing '|' is accepted)."
                )
            if size_mb_column is None and size_bytes_column is None:
                raise ValueError(
                    f"'{filename}' needs a FileSize or FileSizeMB column "
                    "to summarize records that contain data."
                )
            if extension_column is None and filename_column is None:
                raise ValueError(
                    f"'{filename}' needs a FileExtension or FileName column "
                    "to categorize file types."
                )
            reserved = {
                SOURCE_COLUMN,
                MODIFIED_INDEX_COLUMN,
                EXTENSION_INDEX_COLUMN,
                SIZE_MB_COLUMN,
                STATUS_COLUMN,
                SIZE_CATEGORY_COLUMN,
                EXTENSION_CATEGORY_COLUMN,
            }
            collisions = reserved.intersection(columns)
            if collisions:
                raise ValueError(
                    f"'{filename}' uses reserved column name(s): "
                    f"{', '.join(sorted(collisions))}."
                )
        elif set(columns) != set(source_columns):
            missing = [column for column in source_columns if column not in columns]
            extra = [column for column in columns if column not in source_columns]
            raise ValueError(
                f"'{filename}' has different columns. Missing: "
                f"{missing or 'none'}; extra: {extra or 'none'}."
            )
        else:
            frame = frame[source_columns]

        frame[SOURCE_COLUMN] = filename
        frames.append(frame)

    if source_columns is None or modified_column is None:
        raise ValueError("No workbook data was available to analyze.")

    combined = pd.concat(frames, ignore_index=True)
    combined[MODIFIED_INDEX_COLUMN] = pd.to_datetime(
        combined[modified_column], errors="coerce", utc=True
    )

    if size_mb_column is not None:
        size_mb = pd.to_numeric(combined[size_mb_column], errors="coerce")
    else:
        size_mb = pd.Series(float("nan"), index=combined.index, dtype="float64")
    if size_bytes_column is not None:
        size_bytes = pd.to_numeric(combined[size_bytes_column], errors="coerce")
        size_mb = size_mb.fillna(size_bytes / 1_048_576)
    combined[SIZE_MB_COLUMN] = size_mb.where(size_mb >= 0)

    extension_values = (
        combined[extension_column].map(normalize_extension)
        if extension_column is not None
        else pd.Series("", index=combined.index, dtype="string")
    )
    if filename_column is not None:
        derived = combined[filename_column].map(
            lambda value: normalize_extension(
                str(value).rsplit(".", 1)[-1]
                if pd.notna(value) and "." in str(value)
                else ""
            )
        )
        extension_values = extension_values.mask(extension_values.eq(""), derived)
    combined[EXTENSION_INDEX_COLUMN] = extension_values
    combined[SIZE_CATEGORY_COLUMN] = "Size unknown"
    combined.loc[combined[SIZE_MB_COLUMN] > 0, SIZE_CATEGORY_COLUMN] = "Contains data"
    combined.loc[combined[SIZE_MB_COLUMN] == 0, SIZE_CATEGORY_COLUMN] = "0 file size"

    modified_dates = combined[MODIFIED_INDEX_COLUMN]
    latest_modified = modified_dates.max()
    if pd.isna(latest_modified):
        cutoff = None
        activity = pd.Series(
            "Missing modified date", index=combined.index, dtype="string"
        )
    else:
        cutoff = add_months(latest_modified, -6)
        activity = pd.Series(
            "Missing modified date", index=combined.index, dtype="string"
        )
        activity.loc[modified_dates.notna() & (modified_dates >= cutoff)] = "Active"
        activity.loc[modified_dates.notna() & (modified_dates < cutoff)] = "Inactive"
    combined[STATUS_COLUMN] = activity

    all_rows = combined.set_index(
        MODIFIED_INDEX_COLUMN, drop=False
    ).sort_index(kind="stable")
    date_indexed = all_rows.loc[all_rows.index.notna()]
    missing_date_rows = all_rows.loc[all_rows.index.isna()]
    extension_indexed = all_rows.set_index(
        EXTENSION_INDEX_COLUMN, drop=False
    ).sort_index(kind="stable")
    extension_counts = all_rows.groupby(
        EXTENSION_INDEX_COLUMN, dropna=False
    ).size()
    per_file_summary = (
        all_rows.groupby([SOURCE_COLUMN, STATUS_COLUMN], observed=True)
        .size()
        .unstack(fill_value=0)
        .reindex(columns=ACTIVITY_CATEGORIES, fill_value=0)
        .reset_index()
        .rename(columns={SOURCE_COLUMN: "Workbook"})
    )
    per_file_summary.insert(
        1, "Records analyzed", per_file_summary[ACTIVITY_CATEGORIES].sum(axis=1)
    )

    return {
        "source_columns": source_columns,
        "modified_column": modified_column,
        "extension_counts": extension_counts,
        "all_rows": all_rows,
        "date_indexed": date_indexed,
        "missing_date_rows": missing_date_rows,
        "extension_indexed": extension_indexed,
        "latest_modified": latest_modified,
        "cutoff": cutoff,
        "per_file_summary": per_file_summary,
    }


def parse_extension_list(value: str) -> set[str]:
    return {
        normalized
        for token in re.split(r"[\s,;]+", value)
        if (normalized := normalize_extension(token))
    }


def make_pie_chart(data: pd.DataFrame, key: str) -> None:
    chart_data = data.loc[data["Records"] > 0, ["Category", "Records"]]
    if chart_data.empty:
        st.info("There are no records in this chart.")
        return
    total = int(chart_data["Records"].sum())
    angle = 0.0
    labels = []
    for row in chart_data.itertuples(index=False):
        midpoint = angle + (float(row.Records) / total * math.pi)
        labels.append(
            {
                "Category": row.Category,
                "Records": int(row.Records),
                "Label": f"{row.Category}\n{row.Records:,}",
                "X": 0.5 + 0.29 * math.sin(midpoint),
                "Y": 0.5 - 0.29 * math.cos(midpoint),
            }
        )
        angle += float(row.Records) / total * 2 * math.pi
    chart_data = pd.DataFrame(labels)
    st.vega_lite_chart(
        chart_data,
        {
            "layer": [
                {
                    "mark": {
                        "type": "arc",
                        "outerRadius": 105,
                        "tooltip": True,
                    },
                    "encoding": {
                        "theta": {"field": "Records", "type": "quantitative"},
                        "color": {
                            "field": "Category",
                            "type": "nominal",
                            "legend": {"orient": "bottom"},
                        },
                        "tooltip": [
                            {"field": "Category", "type": "nominal"},
                            {"field": "Records", "type": "quantitative"},
                        ],
                    },
                },
                {
                    "mark": {
                        "type": "text",
                        "align": "center",
                        "baseline": "middle",
                        "fontSize": 10,
                        "fontWeight": "bold",
                        "lineBreak": "\n",
                    },
                    "encoding": {
                        "x": {
                            "field": "X",
                            "type": "quantitative",
                            "scale": {"domain": [0, 1]},
                            "axis": None,
                        },
                        "y": {
                            "field": "Y",
                            "type": "quantitative",
                            "scale": {"domain": [0, 1], "range": "height"},
                            "axis": None,
                        },
                        "text": {"field": "Label", "type": "nominal"},
                    },
                },
            ],
            "view": {"stroke": None},
        },
        key=key,
        width="stretch",
        height=285,
    )


def read_keyword_file(uploaded_file: Any) -> list[str]:
    suffix = uploaded_file.name.lower().rsplit(".", 1)[-1]
    contents = uploaded_file.getvalue()
    try:
        if suffix == "txt":
            text = contents.decode("utf-8-sig")
            values = re.split(r"[\r\n,;]+", text)
        elif suffix == "csv":
            frame = pd.read_csv(
                StringIO(contents.decode("utf-8-sig")),
                header=None,
                dtype="string",
            )
            values = frame.to_numpy().ravel().tolist()
        elif suffix in {"xlsx", "xls", "xlsm"}:
            engine = "xlrd" if suffix == "xls" else "openpyxl"
            with pd.ExcelFile(BytesIO(contents), engine=engine) as workbook:
                values = []
                for sheet_name in workbook.sheet_names:
                    frame = pd.read_excel(
                        workbook, sheet_name=sheet_name, header=None, dtype="string"
                    )
                    values.extend(frame.to_numpy().ravel().tolist())
        else:
            raise ValueError("Upload a TXT, CSV, XLSX, XLSM, or XLS keyword list.")
    except (UnicodeDecodeError, ImportError, OSError, ValueError, pd.errors.ParserError) as error:
        raise ValueError(f"Could not read keyword file '{uploaded_file.name}': {error}") from error

    keywords = []
    for value in values:
        if pd.isna(value):
            continue
        keywords.extend(
            token.strip()
            for token in re.split(r"[\r\n,;]+", str(value))
            if token.strip()
        )
    return keywords


def visible_records(frame: pd.DataFrame) -> pd.DataFrame:
    internal_columns = {
        MODIFIED_INDEX_COLUMN,
        EXTENSION_INDEX_COLUMN,
        SIZE_MB_COLUMN,
    }
    return frame.drop(
        columns=[column for column in internal_columns if column in frame.columns]
    )


def prepare_csv_download(
    label: str,
    frame: pd.DataFrame,
    state_key: str,
    filename: str,
) -> None:
    if st.button(f"Prepare {label} CSV", key=f"prepare_{state_key}"):
        st.session_state[state_key] = frame.to_csv(index=False).encode("utf-8-sig")
    csv_data = st.session_state.get(state_key)
    if csv_data is not None:
        st.download_button(
            f"Download {label} CSV",
            data=csv_data,
            file_name=filename,
            mime="text/csv",
            key=f"download_{state_key}",
            on_click="ignore",
        )


st.set_page_config(page_title="Excel Metadata Analyzer", page_icon="📊", layout="wide")
st.title("Excel Metadata Analyzer")
st.write(
    "A clear, business-friendly view of file size, file types, and how recently "
    "your records were modified."
)

uploads = st.file_uploader(
    "Upload one or more Excel inventory files",
    type=["xlsx", "xlsm", "xls"],
    accept_multiple_files=True,
    help=(
        "Every row from every selected worksheet is included. Files are combined "
        "by matching their column names."
    ),
)
if not uploads:
    for state_key in list(st.session_state):
        if state_key.endswith("_indexed_analysis") or state_key.endswith("_sheets"):
            st.session_state.pop(state_key, None)
    st.info("Upload at least one Excel file to create the report.")
    st.stop()

upload_signature = tuple((file.file_id, file.name, file.size) for file in uploads)
cache_key = sha256(repr(upload_signature).encode("utf-8")).hexdigest()
if st.session_state.get("_metadata_upload_signature") != upload_signature:
    for state_key in list(st.session_state):
        if state_key.endswith("_indexed_analysis") or state_key.endswith("_sheets"):
            st.session_state.pop(state_key, None)
    st.session_state["_metadata_upload_signature"] = upload_signature

sheet_state_key = f"{cache_key}_sheets"
file_payloads = None
if sheet_state_key not in st.session_state:
    try:
        file_payloads = tuple((file.name, file.getvalue()) for file in uploads)
        st.session_state[sheet_state_key] = inspect_workbooks(file_payloads)
    except ValueError as error:
        st.error(str(error))
        st.stop()

workbook_sheets = st.session_state[sheet_state_key]
st.subheader("Select the data worksheet in each file")
selected_sheets = []
for index, (uploaded_file, sheet_names) in enumerate(zip(uploads, workbook_sheets)):
    selected_sheets.append(
        st.selectbox(
            f"Worksheet in {uploaded_file.name}",
            options=sheet_names,
            key=f"{cache_key}_sheet_{index}",
        )
    )

analysis_key = (
    f"{cache_key}_{sha256(repr(tuple(selected_sheets)).encode('utf-8')).hexdigest()}"
)
analysis_state_key = f"{analysis_key}_indexed_analysis_v3"
if st.session_state.get(analysis_state_key) is None:
    try:
        if file_payloads is None:
            file_payloads = tuple((file.name, file.getvalue()) for file in uploads)
        with st.spinner("Combining workbooks and building date and extension indexes..."):
            st.session_state[analysis_state_key] = load_and_index_files(
                tuple(
                    (name, contents, sheet_name)
                    for (name, contents), sheet_name in zip(
                        file_payloads, selected_sheets
                    )
                )
            )
    except ValueError as error:
        st.error(str(error))
        st.stop()

analysis = st.session_state[analysis_state_key]
all_rows = analysis["all_rows"]
latest_modified = analysis["latest_modified"]
cutoff = analysis["cutoff"]
source_columns = analysis["source_columns"]

overview_tab, search_tab = st.tabs(["Business summary", "Keyword search"])
with overview_tab:
    st.header("Your data at a glance")
    st.caption(
        "All rows are included, including repeated records. When several files "
        "are uploaded, their selected worksheets are combined."
    )

    count_metrics = st.columns(4)
    count_metrics[0].metric("Files included", f"{len(uploads):,}")
    count_metrics[1].metric("Total records", f"{len(all_rows):,}")
    count_metrics[2].metric(
        "Records with data",
        f"{int((all_rows[SIZE_CATEGORY_COLUMN] == 'Contains data').sum()):,}",
    )
    count_metrics[3].metric(
        "0 file size",
        f"{int((all_rows[SIZE_CATEGORY_COLUMN] == '0 file size').sum()):,}",
    )

    if pd.isna(latest_modified):
        st.warning(
            "No valid Modified dates were found. File size and extension summaries "
            "are available, but activity cannot be determined."
        )
    else:
        st.info(
            f"Activity window: **{cutoff.strftime('%Y-%m-%d')} to "
            f"{latest_modified.strftime('%Y-%m-%d')}** (inclusive), based on the "
            "latest valid Modified date in all selected files."
        )

    st.subheader("The three business questions")
    size_counts = (
        all_rows[SIZE_CATEGORY_COLUMN]
        .value_counts()
        .reindex(SIZE_CATEGORIES, fill_value=0)
        .rename_axis("Category")
        .reset_index(name="Records")
    )
    activity_counts = (
        all_rows[STATUS_COLUMN]
        .value_counts()
        .reindex(ACTIVITY_CATEGORIES, fill_value=0)
        .rename_axis("Category")
        .reset_index(name="Records")
    )

    extension_map = {}
    for category, configured_extensions in DEFAULT_EXTENSIONS.items():
        for extension in configured_extensions:
            extension_map[extension] = category

    with st.expander("Customize file-extension buckets", expanded=False):
        st.write(
            "Enter extensions with or without a leading dot. Separate entries "
            "with commas, spaces, or new lines. Any extension not listed is "
            "Undetermined."
        )
        st.info(
            "NIST's [National Software Reference Library (NSRL)]("
            f"{NIST_NSRL_URL}) is a repository of software file signatures, "
            "not an official ROT-by-extension list. File extensions alone cannot "
            "verify an NSRL signature. ROT defaults below are a transparent, "
            "editable starting point - not a NIST endorsement."
        )
        configured = {}
        for category in EXTENSION_CATEGORIES[:3]:
            initial = ", ".join(DEFAULT_EXTENSIONS[category])
            configured[category] = st.text_area(
                category,
                value=initial,
                key=f"{analysis_key}_extensions_{category}",
                height=110,
            )
        extension_map = {}
        overlaps = {}
        for category, value in configured.items():
            for extension in parse_extension_list(value):
                if extension in extension_map:
                    overlaps.setdefault(extension, [extension_map[extension]]).append(
                        category
                    )
                else:
                    extension_map[extension] = category
        if overlaps:
            overlap_description = "; ".join(
                f"{extension}: {', '.join(categories)}"
                for extension, categories in sorted(overlaps.items())
            )
            st.error(
                "Each extension can belong to only one bucket. Resolve these "
                f"overlaps to update the report: {overlap_description}"
            )
            st.stop()

    extension_categories = all_rows[EXTENSION_INDEX_COLUMN].map(
        lambda extension: extension_map.get(extension, "Undetermined files")
    )
    all_rows[EXTENSION_CATEGORY_COLUMN] = extension_categories
    analysis["extension_indexed"][EXTENSION_CATEGORY_COLUMN] = analysis[
        "extension_indexed"
    ][EXTENSION_INDEX_COLUMN].map(
        lambda extension: extension_map.get(extension, "Undetermined files")
    )
    extension_counts = (
        extension_categories.value_counts()
        .reindex(EXTENSION_CATEGORIES, fill_value=0)
        .rename_axis("Category")
        .reset_index(name="Records")
    )

    chart_columns = st.columns(3)
    with chart_columns[0]:
        st.markdown("**File-size status**")
        make_pie_chart(size_counts, f"{analysis_key}_size_pie")
        st.dataframe(size_counts, width="stretch", hide_index=True)
    with chart_columns[1]:
        st.markdown("**File types**")
        make_pie_chart(extension_counts, f"{analysis_key}_extension_pie")
        st.dataframe(extension_counts, width="stretch", hide_index=True)
    with chart_columns[2]:
        st.markdown("**Activity**")
        make_pie_chart(activity_counts, f"{analysis_key}_activity_pie")
        st.dataframe(activity_counts, width="stretch", hide_index=True        )

        st.subheader("Extensions included in each bucket")
        observed_extension_counts = analysis["extension_counts"].to_dict()
        extensions_to_report = set(extension_map).union(observed_extension_counts)
        extension_summary = pd.DataFrame(
            {
                "File extension": sorted(extensions_to_report),
            }
        )
        extension_summary["Business bucket"] = extension_summary[
            "File extension"
        ].map(lambda extension: extension_map.get(extension, "Undetermined files"))
        extension_summary["Records"] = extension_summary["File extension"].map(
            observed_extension_counts
        ).fillna(0).astype("int64")
        extension_summary["File extension"] = extension_summary[
            "File extension"
        ].replace("", "(no extension)")
    extension_summary = extension_summary.sort_values(
        ["Business bucket", "Records", "File extension"],
        ascending=[True, False, True],
        kind="stable",
    )
    st.dataframe(
        extension_summary,
        width="stretch",
        hide_index=True,
        height=300,
    )
    st.caption(
        "The table shows observed extensions and their current bucket. Edit the "
        "bucket lists above to reflect your organization’s business decisions."
    )

    st.subheader("Explore matching records")
    filter_columns = st.columns(3)
    with filter_columns[0]:
        selected_extension_category = st.selectbox(
            "File-type bucket",
            ["All buckets", *EXTENSION_CATEGORIES],
            key=f"{analysis_key}_extension_filter",
        )
    with filter_columns[1]:
        selected_size_category = st.selectbox(
            "File-size status",
            ["All sizes", *SIZE_CATEGORIES],
            key=f"{analysis_key}_size_filter",
        )
    with filter_columns[2]:
        selected_activity = st.selectbox(
            "Activity status",
            ["All activity", *ACTIVITY_CATEGORIES],
            key=f"{analysis_key}_activity_filter",
        )

    filtered = all_rows
    if selected_extension_category != "All buckets":
        matching_extensions = [
            extension
            for extension in analysis["extension_counts"].index
            if extension_map.get(extension, "Undetermined files")
            == selected_extension_category
        ]
        if matching_extensions:
            filtered = analysis["extension_indexed"].loc[matching_extensions]
        else:
            filtered = all_rows.iloc[0:0]
    if selected_size_category != "All sizes":
        filtered = filtered.loc[
            filtered[SIZE_CATEGORY_COLUMN] == selected_size_category
        ]
    if selected_activity != "All activity":
        filtered = filtered.loc[filtered[STATUS_COLUMN] == selected_activity]

    st.markdown(f"**Matching records: {len(filtered):,}**")
    preview_limit = st.selectbox(
        "Rows to preview",
        [100, 500, 1000, 5000],
        index=2,
        key=f"{analysis_key}_preview_limit",
    )
    st.dataframe(
        visible_records(filtered).head(preview_limit),
        width="stretch",
        height=420,
        hide_index=True,
    )
    prepare_csv_download(
        "filtered records",
        visible_records(filtered),
        f"{analysis_key}_overview_csv",
        "filtered_inventory_records.csv",
    )

    with st.expander("Record counts by workbook"):
        st.dataframe(
            analysis["per_file_summary"],
            width="stretch",
            hide_index=True,
        )

with search_tab:
    st.header("Find records by keywords")
    st.write(
        "Choose a source column and search for one or more words. Matching is "
        "case-insensitive and returns rows containing **any** supplied keyword."
    )
    search_column = st.selectbox(
        "Column to search",
        source_columns,
        format_func=str,
        key=f"{analysis_key}_search_column",
    )
    typed_keywords = st.text_area(
        "Type keywords (separate with commas, semicolons, or new lines)",
        key=f"{analysis_key}_typed_keywords",
        height=120,
    )
    keyword_file = st.file_uploader(
        "Or upload a keyword list",
        type=["txt", "csv", "xlsx", "xlsm", "xls"],
        accept_multiple_files=False,
        key=f"{analysis_key}_keyword_file",
        help="Text and spreadsheet lists are read from all cells and sheets.",
    )
    typed_terms = [
        token.strip()
        for token in re.split(r"[\r\n,;]+", typed_keywords)
        if token.strip()
    ]
    uploaded_terms = []
    if keyword_file is not None:
        try:
            uploaded_terms = read_keyword_file(keyword_file)
        except ValueError as error:
            st.error(str(error))
            st.stop()
    keywords = list(dict.fromkeys(typed_terms + uploaded_terms))

    if st.button("Search records", key=f"{analysis_key}_run_search"):
        if not keywords:
            st.session_state.pop(f"{analysis_key}_search_results", None)
            st.warning("Enter or upload at least one keyword before searching.")
        else:
            alternatives = "|".join(re.escape(keyword) for keyword in keywords)
            matching_mask = (
                all_rows[search_column]
                .astype("string")
                .str.contains(alternatives, case=False, regex=True, na=False)
            )
            st.session_state[f"{analysis_key}_search_results"] = {
                "column": search_column,
                "keywords": keywords,
                "rows": all_rows.loc[matching_mask],
            }

    saved_search = st.session_state.get(f"{analysis_key}_search_results")
    if saved_search is None:
        st.info(
            "Enter keywords (or upload a list), then select Search records. "
            "A keyword matches when it appears anywhere within the selected "
            "column value."
        )
    else:
        search_results = saved_search["rows"]
        st.subheader("Search results")
        st.caption(
            f"Column: {saved_search['column']} | "
            f"Keywords: {', '.join(saved_search['keywords'])}"
        )
        st.markdown(
            f"**{len(search_results):,} matching record(s)** - results include "
            "values containing any of the supplied keywords."
        )
        if search_results.empty:
            st.info("No records matched these keywords.")
        else:
            st.dataframe(
                visible_records(search_results).head(500),
                width="stretch",
                height=450,
                hide_index=True,
            )
            if len(search_results) > 500:
                st.caption(
                    "Showing the first 500 matches. The CSV contains all matches."
                )
            prepare_csv_download(
                "search results",
                visible_records(search_results),
                f"{analysis_key}_search_csv",
                "metadata_search_results.csv",
            )
