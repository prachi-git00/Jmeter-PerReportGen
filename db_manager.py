import os
import sqlite3
from contextlib import contextmanager
from datetime import datetime

DB_FILE = os.environ.get("DB_FILE", "pretest_changes.db")
DATABASE_URL = os.environ.get("DATABASE_URL") or os.environ.get("POSTGRES_URL")


def _is_postgres_enabled():
    """Return True when PostgreSQL is configured for deployment."""
    db_type = os.environ.get("DB_TYPE", "sqlite").lower()
    return db_type in {"postgres", "postgresql"} or bool(DATABASE_URL)


def _prepare_sql(sql):
    """Convert SQLite placeholder syntax to PostgreSQL syntax when needed."""
    if _is_postgres_enabled():
        return sql.replace("?", "%s")
    return sql


@contextmanager
def _connection():
    if _is_postgres_enabled():
        try:
            import psycopg2
        except ModuleNotFoundError as exc:
            raise RuntimeError(
                "psycopg2-binary is required for PostgreSQL support. "
                "Install it in requirements.txt or set DATABASE_URL for SQLite fallback."
            ) from exc

        conn = psycopg2.connect(DATABASE_URL, connect_timeout=10)
        try:
            yield conn
        finally:
            conn.close()
        return

    conn = sqlite3.connect(DB_FILE, timeout=10.0)
    try:
        yield conn
    finally:
        conn.close()


def _table_columns(conn):
    """Return column names for the given table across database backends."""
    if _is_postgres_enabled():
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_name = %s
            ORDER BY ordinal_position
            """,
            ("pretest_changes",),
        )
        return [row[0] for row in cursor.fetchall()]

    cursor = conn.cursor()
    cursor.execute('PRAGMA table_info(pretest_changes)')
    return [row[1] for row in cursor.fetchall()]


def _duplicate_error(exc):
    """Detect duplicate-key conflicts across SQLite and PostgreSQL."""
    message = str(exc).lower()
    return (
        "integrity" in message
        or "duplicate key" in message
        or "unique constraint" in message
        or "already exists" in message
    )


def init_database():
    """Initialize the database with required tables."""
    with _connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            _prepare_sql(
                '''
                CREATE TABLE IF NOT EXISTS projects (
                    project_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_name TEXT UNIQUE NOT NULL,
                    created_date TIMESTAMP
                )
                '''
            )
        )
        cursor.execute(
            _prepare_sql(
                '''
                CREATE TABLE IF NOT EXISTS pretest_changes (
                    change_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id INTEGER NOT NULL,
                    run_id TEXT NOT NULL,
                    tier TEXT NOT NULL,
                    changes TEXT NOT NULL,
                    change_date DATE NOT NULL,
                    created_timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (project_id) REFERENCES projects(project_id)
                )
                '''
            )
        )

        columns = _table_columns(conn)
        if 'tier' not in columns:
            if _is_postgres_enabled():
                cursor.execute("ALTER TABLE pretest_changes ADD COLUMN tier TEXT DEFAULT ''")
            else:
                cursor.execute('ALTER TABLE pretest_changes ADD COLUMN tier TEXT DEFAULT ""')

        cursor.execute(
            _prepare_sql(
                '''
                CREATE INDEX IF NOT EXISTS idx_pretest_changes_project_date
                ON pretest_changes (project_id, change_date DESC, created_timestamp DESC)
                '''
            )
        )


def get_all_projects():
    """Get list of all projects."""
    with _connection() as conn:
        cursor = conn.execute(_prepare_sql('SELECT project_name FROM projects ORDER BY project_name'))
        return [row[0] for row in cursor.fetchall()]


def create_project(project_name):
    """Create a new project."""
    with _connection() as conn:
        try:
            conn.execute(
                _prepare_sql('INSERT INTO projects (project_name) VALUES (?)'),
                (project_name,),
            )
            conn.commit()
            return True
        except Exception as exc:
            if _duplicate_error(exc):
                return False
            raise


def get_project_id(project_name):
    """Get project ID by project name."""
    with _connection() as conn:
        result = conn.execute(
            _prepare_sql('SELECT project_id FROM projects WHERE project_name = ?'),
            (project_name,),
        ).fetchone()
        return result[0] if result else None


def add_pretest_change(project_name, run_id, tier, changes, change_date):
    """Add a new pretest change for a project."""
    with _connection() as conn:
        project = conn.execute(
            _prepare_sql('SELECT project_id FROM projects WHERE project_name = ?'),
            (project_name,),
        ).fetchone()
        if not project:
            return False, None
        project_id = project[0]
        is_duplicate = conn.execute(
            _prepare_sql('''
                SELECT 1 FROM pretest_changes
                WHERE project_id = ? AND run_id = ?
            '''),
            (project_id, run_id),
        ).fetchone() is not None
        conn.execute(
            _prepare_sql('''
                INSERT INTO pretest_changes (project_id, run_id, tier, changes, change_date)
                VALUES (?, ?, ?, ?, ?)
            '''),
            (project_id, run_id, tier, changes, change_date),
        )
        conn.commit()
        return True, is_duplicate


def update_pretest_change(project_name, run_id, tier, changes, change_date):
    """Update an existing pretest change identified by project and run_id."""
    with _connection() as conn:
        cursor = conn.execute(
            _prepare_sql('''
                UPDATE pretest_changes
                SET tier = ?, changes = ?, change_date = ?
                WHERE project_id = (
                    SELECT project_id FROM projects WHERE project_name = ?
                ) AND run_id = ?
            '''),
            (tier, changes, change_date, project_name, run_id),
        )
        conn.commit()
        return cursor.rowcount > 0


def _parse_change_date(date_value):
    """Return a comparable date value for pretest-change records."""
    if date_value is None:
        return datetime.min.date()

    date_str = str(date_value).strip()
    if not date_str:
        return datetime.min.date()

    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y", "%d-%m-%Y %H:%M:%S"):
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(date_str).date()
    except ValueError:
        return datetime.min.date()


def get_pretest_changes(project_name):
    """Get all pretest changes for a project, sorted by date (newest first)."""
    return get_pretest_changes_filtered(project_name)


def get_pretest_changes_filtered(project_name, run_id=None, tier=None, start_date=None, end_date=None):
    """Get pretest changes for a project filtered by Run ID, Tier, and/or date range."""
    query = '''
        SELECT run_id, tier, changes, change_date
        FROM pretest_changes
        WHERE project_id = (
            SELECT project_id FROM projects WHERE project_name = ?
        )
    '''
    params = [project_name]

    if run_id:
        query += ' AND run_id LIKE ?'
        params.append(f'%{run_id}%')
    if tier:
        query += ' AND tier LIKE ?'
        params.append(f'%{tier}%')
    if start_date:
        query += ' AND change_date >= ?'
        params.append(start_date)
    if end_date:
        query += ' AND change_date <= ?'
        params.append(end_date)

    query += ' ORDER BY date(change_date) DESC, created_timestamp DESC'

    with _connection() as conn:
        return conn.execute(_prepare_sql(query), tuple(params)).fetchall()


def get_pretest_changes_as_dataframe(project_name):
    """Get pretest changes as a pandas DataFrame with Sr. No. column."""
    import pandas as pd

    changes = get_pretest_changes(project_name)

    if not changes:
        return pd.DataFrame(columns=['Sr. No.', 'Run ID', 'Tier', 'Changes Done', 'Date'])

    data = []
    for idx, (run_id, tier, changes_text, change_date) in enumerate(changes, 1):
        data.append({
            'Sr. No.': idx,
            'Run ID': run_id,
            'Tier': tier,
            'Changes Done': changes_text,
            'Date': change_date,
        })

    return pd.DataFrame(data)


def delete_pretest_change(project_name, run_id):
    """Delete a pretest change by run_id."""
    with _connection() as conn:
        cursor = conn.execute(
            _prepare_sql('''
                DELETE FROM pretest_changes
                WHERE project_id = (
                    SELECT project_id FROM projects WHERE project_name = ?
                ) AND run_id = ?
            '''),
            (project_name, run_id),
        )
        conn.commit()
        return cursor.rowcount > 0


def delete_project(project_name):
    """Delete a project and all of its pretest changes."""
    with _connection() as conn:
        project = conn.execute(
            _prepare_sql('SELECT project_id FROM projects WHERE project_name = ?'),
            (project_name,),
        ).fetchone()
        if not project:
            return False
        project_id = project[0]
        conn.execute(
            _prepare_sql('DELETE FROM pretest_changes WHERE project_id = ?'),
            (project_id,),
        )
        conn.execute(
            _prepare_sql('DELETE FROM projects WHERE project_id = ?'),
            (project_id,),
        )
        conn.commit()
        return True
