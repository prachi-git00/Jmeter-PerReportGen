import sqlite3
from contextlib import contextmanager
from datetime import datetime

# Database file location
DB_FILE = "pretest_changes.db"


@contextmanager
def _connection():
    conn = sqlite3.connect(DB_FILE, timeout=10.0)
    try:
        yield conn
    finally:
        conn.close()


def init_database():
    """Initialize the database with required tables"""
    with _connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS projects (
                project_id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_name TEXT UNIQUE NOT NULL,
                created_date TIMESTAMP
            )
        ''')
        cursor.execute('''
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
        ''')
        cursor.execute('PRAGMA table_info(pretest_changes)')
        columns = [row[1] for row in cursor.fetchall()]
        if 'tier' not in columns:
            cursor.execute('ALTER TABLE pretest_changes ADD COLUMN tier TEXT DEFAULT ""')
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_pretest_changes_project_date
            ON pretest_changes (project_id, change_date DESC, created_timestamp DESC)
        ''')

def get_all_projects():
    """Get list of all projects"""
    with _connection() as conn:
        cursor = conn.execute('SELECT project_name FROM projects ORDER BY project_name')
        return [row[0] for row in cursor.fetchall()]

def create_project(project_name):
    """Create a new project"""
    with _connection() as conn:
        try:
            conn.execute('INSERT INTO projects (project_name) VALUES (?)', (project_name,))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False  # Project already exists

def get_project_id(project_name):
    """Get project ID by project name"""
    with _connection() as conn:
        result = conn.execute(
            'SELECT project_id FROM projects WHERE project_name = ?', (project_name,)
        ).fetchone()
        return result[0] if result else None

def add_pretest_change(project_name, run_id, tier, changes, change_date):
    """Add a new pretest change for a project"""
    with _connection() as conn:
        project = conn.execute(
            'SELECT project_id FROM projects WHERE project_name = ?', (project_name,)
        ).fetchone()
        if not project:
            return False, None
        project_id = project[0]
        is_duplicate = conn.execute('''
            SELECT 1 FROM pretest_changes
            WHERE project_id = ? AND run_id = ?
        ''', (project_id, run_id)).fetchone() is not None
        conn.execute('''
            INSERT INTO pretest_changes (project_id, run_id, tier, changes, change_date)
            VALUES (?, ?, ?, ?, ?)
        ''', (project_id, run_id, tier, changes, change_date))
        conn.commit()
        return True, is_duplicate

def update_pretest_change(project_name, run_id, tier, changes, change_date):
    """Update an existing pretest change identified by project and run_id"""
    with _connection() as conn:
        cursor = conn.execute('''
            UPDATE pretest_changes
            SET tier = ?, changes = ?, change_date = ?
            WHERE project_id = (
                SELECT project_id FROM projects WHERE project_name = ?
            ) AND run_id = ?
        ''', (tier, changes, change_date, project_name, run_id))
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
    """Get all pretest changes for a project, sorted by date (newest first)"""
    return get_pretest_changes_filtered(project_name)

def get_pretest_changes_filtered(project_name, run_id=None, tier=None, start_date=None, end_date=None):
    """Get pretest changes for a project filtered by Run ID, Tier, and/or date range"""
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
        return conn.execute(query, tuple(params)).fetchall()

def get_pretest_changes_as_dataframe(project_name):
    """Get pretest changes as a pandas DataFrame with Sr. No. column"""
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
            'Date': change_date
        })
    
    return pd.DataFrame(data)

def delete_pretest_change(project_name, run_id):
    """Delete a pretest change by run_id"""
    with _connection() as conn:
        cursor = conn.execute('''
            DELETE FROM pretest_changes
            WHERE project_id = (
                SELECT project_id FROM projects WHERE project_name = ?
            ) AND run_id = ?
        ''', (project_name, run_id))
        conn.commit()
        return cursor.rowcount > 0

def delete_project(project_name):
    """Delete a project and all of its pretest changes"""
    with _connection() as conn:
        project = conn.execute(
            'SELECT project_id FROM projects WHERE project_name = ?', (project_name,)
        ).fetchone()
        if not project:
            return False
        project_id = project[0]
        conn.execute('DELETE FROM pretest_changes WHERE project_id = ?', (project_id,))
        conn.execute('DELETE FROM projects WHERE project_id = ?', (project_id,))
        conn.commit()
        return True
