from __future__ import annotations

import base64
from typing import Any

from psycopg import sql

from backend.db import get_connection
from backend.schema_map import DOC_SHEETS, get_columns, get_headers, get_table_name


def _select_columns(sheet_name: str):
    headers = get_headers(sheet_name)
    columns = get_columns(sheet_name)
    return sql.SQL(", ").join(
        sql.SQL("{column} AS {alias}").format(
            column=sql.Identifier(column),
            alias=sql.Identifier(header),
        )
        for header, column in zip(headers, columns)
    )


def list_records(sheet_name: str, filters: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    table_name = get_table_name(sheet_name)
    headers = get_headers(sheet_name)
    columns = get_columns(sheet_name)
    primary_key_column = columns[0]
    query = sql.SQL("SELECT {cols} FROM {table}").format(
        cols=_select_columns(sheet_name),
        table=sql.Identifier(table_name),
    )

    params: list[Any] = []
    if filters:
        filter_parts = []
        for header, value in filters.items():
            if header not in headers:
                continue
            column = columns[headers.index(header)]
            filter_parts.append(sql.SQL("{column} = {value}").format(column=sql.Identifier(column), value=sql.Placeholder()))
            params.append(value)
        if filter_parts:
            query += sql.SQL(" WHERE ") + sql.SQL(" AND ").join(filter_parts)

    query += sql.SQL(" ORDER BY {pk}").format(pk=sql.Identifier(primary_key_column))

    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(query, params)
            rows = cur.fetchall()
    return [dict(row) for row in rows]


def get_record_by_key(sheet_name: str, key_header: str, key_value: Any) -> dict[str, Any] | None:
    rows = list_records(sheet_name, {key_header: key_value})
    return rows[0] if rows else None


def append_record(sheet_name: str, row_dict: dict[str, Any]) -> None:
    headers = get_headers(sheet_name)
    columns = get_columns(sheet_name)
    values = [row_dict.get(header) for header in headers]
    query = sql.SQL("INSERT INTO {table} ({cols}) VALUES ({vals})").format(
        table=sql.Identifier(get_table_name(sheet_name)),
        cols=sql.SQL(", ").join(sql.Identifier(col) for col in columns),
        vals=sql.SQL(", ").join(sql.Placeholder() for _ in columns),
    )
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(query, values)


def update_record_by_key(sheet_name: str, key_header: str, key_value: Any, updates: dict[str, Any]) -> bool:
    headers = get_headers(sheet_name)
    columns = get_columns(sheet_name)
    key_index = headers.index(key_header)
    key_column = columns[key_index]

    filtered_updates = {header: value for header, value in updates.items() if header in headers}
    if not filtered_updates:
        return False

    query = sql.SQL("UPDATE {table} SET {updates} WHERE {key_col} = {key_value}").format(
        table=sql.Identifier(get_table_name(sheet_name)),
        updates=sql.SQL(", ").join(
            sql.SQL("{col} = {placeholder}").format(col=sql.Identifier(columns[headers.index(header)]), placeholder=sql.Placeholder())
            for header in filtered_updates
        ),
        key_col=sql.Identifier(key_column),
        key_value=sql.Placeholder(),
    )
    params = list(filtered_updates.values()) + [key_value]
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(query, params)
            return cur.rowcount > 0


def upsert_record(sheet_name: str, key_header: str, row_dict: dict[str, Any]) -> str:
    headers = get_headers(sheet_name)
    columns = get_columns(sheet_name)
    key_index = headers.index(key_header)
    key_column = columns[key_index]
    values = [row_dict.get(header) for header in headers]
    existed = get_record_by_key(sheet_name, key_header, row_dict.get(key_header)) is not None

    query = sql.SQL(
        "INSERT INTO {table} ({cols}) VALUES ({vals}) "
        "ON CONFLICT ({key_col}) DO UPDATE SET {updates}"
    ).format(
        table=sql.Identifier(get_table_name(sheet_name)),
        cols=sql.SQL(", ").join(sql.Identifier(col) for col in columns),
        vals=sql.SQL(", ").join(sql.Placeholder() for _ in columns),
        key_col=sql.Identifier(key_column),
        updates=sql.SQL(", ").join(
            sql.SQL("{col} = EXCLUDED.{col}").format(col=sql.Identifier(col))
            for idx, col in enumerate(columns)
            if idx != key_index
        ),
    )
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(query, values)
    return "updated" if existed else "inserted"


def delete_record_by_key(sheet_name: str, key_header: str, key_value: Any) -> bool:
    headers = get_headers(sheet_name)
    columns = get_columns(sheet_name)
    key_column = columns[headers.index(key_header)]
    query = sql.SQL("DELETE FROM {table} WHERE {key_col} = {key_value}").format(
        table=sql.Identifier(get_table_name(sheet_name)),
        key_col=sql.Identifier(key_column),
        key_value=sql.Placeholder(),
    )
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(query, [key_value])
            return cur.rowcount > 0


def clear_records(sheet_name: str) -> None:
    query = sql.SQL("TRUNCATE TABLE {table}").format(table=sql.Identifier(get_table_name(sheet_name)))
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(query)


def rewrite_records(sheet_name: str, rows: list[dict[str, Any]]) -> None:
    clear_records(sheet_name)
    for row in rows:
        append_record(sheet_name, row)


def save_attachment(
    sheet_name: str,
    record_id: str,
    file_name: str,
    original_path: str,
    content_type: str,
    content_base64: str,
    created_on: str | None = None,
    updated_on: str | None = None,
) -> str:
    data = base64.b64decode(content_base64.encode("ascii"))
    query = """
        INSERT INTO document_blobs
            (sheet_name, record_id, file_name, original_path, content_type, file_data, file_size, created_on, updated_on)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (sheet_name, record_id)
        DO UPDATE SET
            file_name = EXCLUDED.file_name,
            original_path = EXCLUDED.original_path,
            content_type = EXCLUDED.content_type,
            file_data = EXCLUDED.file_data,
            file_size = EXCLUDED.file_size,
            updated_on = EXCLUDED.updated_on
    """
    timestamp = updated_on or created_on
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                query,
                [sheet_name, record_id, file_name, original_path, content_type, data, len(data), created_on or timestamp, timestamp],
            )
    return f"db://{sheet_name}/{record_id}/{file_name}"


def delete_attachment(sheet_name: str, record_id: str) -> None:
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM document_blobs WHERE sheet_name = %s AND record_id = %s", [sheet_name, record_id])


def get_attachment(sheet_name: str, record_id: str) -> dict[str, Any] | None:
    if sheet_name not in DOC_SHEETS:
        raise ValueError(f"Attachments are not supported for sheet: {sheet_name}")

    query = """
        SELECT file_name, content_type, file_data, file_size, created_on, updated_on
        FROM document_blobs
        WHERE sheet_name = %s AND record_id = %s
    """
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(query, [sheet_name, record_id])
            row = cur.fetchone()
    return dict(row) if row else None
