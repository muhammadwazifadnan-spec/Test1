import os
import re
import pyodbc
from datetime import datetime
from typing import List, Dict, Optional

print("DB FILE:", __file__)

class SQLServerCursorWrapper:
    def __init__(self, conn, dictionary=False):
        self.conn = conn
        self.cur = conn.raw_cursor()
        self.dictionary = dictionary

    def execute(self, query, params=None):
        params = params or []
        query = re.sub(r"%s", "?", query)
        self.cur.execute(query, params)
        return self

    def fetchone(self):
        row = self.cur.fetchone()
        if row is None:
            return None
        if self.dictionary:
            cols = [c[0] for c in self.cur.description]
            return dict(zip(cols, row))
        return row

    def fetchall(self):
        rows = self.cur.fetchall()
        if self.dictionary:
            cols = [c[0] for c in self.cur.description]
            return [dict(zip(cols, row)) for row in rows]
        return rows

    @property
    def description(self):
        return self.cur.description

    @property
    def lastrowid(self):
        try:
            self.cur.execute("SELECT SCOPE_IDENTITY()")
            row = self.cur.fetchone()
            return row[0] if row else None
        except Exception:
            return None

    def close(self):
        self.cur.close()


class SQLServerConnectionWrapper:
    def __init__(self, conn):
        self.conn = conn

    def cursor(self, dictionary=False):
        return SQLServerCursorWrapper(self, dictionary=dictionary)

    def raw_cursor(self):
        return self.conn.cursor()

    def commit(self):
        self.conn.commit()

    def close(self):
        self.conn.close()


def get_connection():
    conn_str = (
        f"DRIVER={{{os.getenv('SQLSERVER_DRIVER')}}};"
        f"SERVER={os.getenv('SQLSERVER_SERVER')};"
        f"DATABASE={os.getenv('SQLSERVER_DATABASE')};"
        "Trusted_Connection=yes;"
        f"TrustServerCertificate={os.getenv('SQLSERVER_TRUST_CERTIFICATE', 'yes')};"
    )
    raw_conn = pyodbc.connect(conn_str)
    return SQLServerConnectionWrapper(raw_conn)

    

def get_db_connection():
    return get_connection()

def _rows_to_dict_list(cursor) -> List[Dict]:
    cols = [c[0] for c in cursor.description]
    return [dict(zip(cols, row)) for row in cursor.fetchall()]


def insert_session_record(
    *,
    file_name: str,
    audio_path: str,
    file_type: str,
    transcript: str,
    translation: str = None,
    sentiment_label: str = None,
    sentiment_score = None, 
    sentiment_tone: str = None,
    explanation: str = None,
    scenario_id: int = None,
    language_used: str = "Unknown",
    file_created_at=None,
    uploaded_at=None
):
    connection = get_db_connection()
    if not connection:
        return
    cursor = connection.cursor()
    sql = """
    INSERT INTO audio_sessions (
        audio_filename,
        audio_path,
        file_type,
        transcript_raw,
        transcript_english,
        sentiment_label,
        sentiment_score,
        sentiment_tone,
        sentiment_explanation,
        scenario_id,
        language_used,
        file_created_at,
        uploaded_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    values = (
        file_name,
        audio_path,
        file_type,
        transcript,
        translation,
        sentiment_label,
        sentiment_score,
        sentiment_tone,
        explanation,
        scenario_id,
        language_used,
        file_created_at,
        uploaded_at
    )
    try:
        cursor.execute(sql, values)
        connection.commit()
        print(f"[DB] Inserted AUDIO session: {file_name} (type={file_type})")
    except Exception as e:
        print("[DB ERROR]", e)
    finally:
        cursor.close()
        connection.close()


def insert_text_record(
    *,
    file_name: str,
    text_path: str,
    file_type: str,
    transcript: str,
    translation: str = None,
    sentiment_label: str = None,
    sentiment_score = None,
    sentiment_tone: str = None,
    explanation: str = None,
    scenario_id: int = None,
    language_used: str = "Unknown",
    file_created_at=None,
    uploaded_at=None
):
    connection = get_db_connection()
    if not connection:
        return
    cursor = connection.cursor()
    sql = """
    INSERT INTO text_sessions (
        text_filename,
        text_path,
        file_type,
        transcript_raw,
        transcript_english,
        sentiment_label,
        sentiment_score,
        sentiment_tone,
        sentiment_explanation,
        scenario_id,
        language_used,
        file_created_at,
        uploaded_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    values = (
        file_name,
        text_path,
        file_type,
        transcript,
        translation,
        sentiment_label,
        sentiment_score,
        sentiment_tone,
        explanation,
        scenario_id,
        language_used,
        file_created_at,
        uploaded_at
    )
    try:
        cursor.execute(sql, values)
        connection.commit()
        print(f"[DB] Inserted TEXT session: {file_name} (type={file_type})")
    except Exception as e:
        print("[DB ERROR]", e)
    finally:
        cursor.close()
        connection.close()


def get_all_scenarios() -> List[Dict]:
    connection = get_db_connection()
    if not connection:
        return []
    cursor = connection.cursor()
    try:
        cursor.execute("""
            SELECT
                scenario_id       AS id,
                scenario_name     AS name,
                scenario_description AS description
            FROM scenarios
        """)
        return _rows_to_dict_list(cursor)
    except Exception as e:
        print("[DB ERROR]", e)
        return []
    finally:
        cursor.close()
        connection.close()

def update_human_sentiment_label(session_id: int, human_label: str) -> bool:
    """
    Store CS correction label (Complaint / Non-Complaint) for a given audio_sessions row.
    """
    connection = get_db_connection()
    if not connection:
        return False
    cursor = connection.cursor()
    try:
        cursor.execute("""
            UPDATE audio_sessions
            SET human_sentiment_label = ?,
                human_updated_at      = ?
            WHERE session_id = ?
        """, (human_label, datetime.now(), int(session_id)))
        connection.commit()
        return True
    except Exception as e:
        print("[DB ERROR]", e)
        return False
    finally:
        cursor.close()
        connection.close()

def fetch_sessions_for_ui(limit: int = 500) -> List[Dict]:
    """
    Returns combined latest sessions (audio + text) for UI.
    """
    connection = get_db_connection()
    if not connection:
        return []
    cursor = connection.cursor()
    try:
        sql = """
        SELECT
            u.source_type,
            u.session_pk,
            u.file_name,
            u.file_type,
            u.transcript_raw,
            u.transcript_english,
            u.sentiment_label,
            u.sentiment_score,
            u.sentiment_tone,
            u.sentiment_explanation,
            u.scenario_id,
            u.uploaded_at,
            u.human_sentiment_label,
            u.human_updated_at
        FROM (
            SELECT
                'audio' AS source_type,
                a.session_id        AS session_pk,
                a.audio_filename    AS file_name,
                a.file_type,
                a.transcript_raw,
                a.transcript_english,
                a.sentiment_label,
                CAST(a.sentiment_score AS DECIMAL(10,2)) AS sentiment_score,
                a.sentiment_tone,
                a.sentiment_explanation,
                a.scenario_id,
                a.uploaded_at,
                a.human_sentiment_label,
                a.human_updated_at
            FROM audio_sessions a

            UNION ALL

            SELECT
                'text' AS source_type,
                t.id              AS session_pk,
                t.text_filename   AS file_name,
                t.file_type,
                t.transcript_raw,
                t.transcript_english,
                t.sentiment_label,
                CAST(t.sentiment_score AS DECIMAL(10,2)) AS sentiment_score,
                t.sentiment_tone,
                t.sentiment_explanation,
                t.scenario_id,
                t.uploaded_at,
                t.human_sentiment_label,
                t.human_updated_at
            FROM text_sessions t
        ) AS u
        ORDER BY u.uploaded_at DESC
        OFFSET 0 ROWS FETCH NEXT ? ROWS ONLY;
        """
        cursor.execute(sql, (int(limit),))
        return _rows_to_dict_list(cursor)
    except Exception as e:
        print("[DB ERROR]", e)
        return []
    finally:
        cursor.close()
        connection.close()


def find_admin(admin_username: str) -> Optional[Dict]:
    connection = get_db_connection()
    if not connection:
        return None
    cursor = connection.cursor()
    try:
        cursor.execute("""
            SELECT TOP 1
                adminID,
                admin_username,
                admin_password
            FROM admin_account
            WHERE admin_username = ?
        """, (admin_username,))
        row = cursor.fetchone()
        if not row:
            return None
        cols = [c[0] for c in cursor.description]
        return dict(zip(cols, row))
    except Exception as e:
        print("[DB ERROR]", e)
        return None
    finally:
        cursor.close()
        connection.close()

def find_user(username: str) -> Optional[Dict]:
    connection = get_db_connection()
    if not connection:
        return None
    cursor = connection.cursor()
    try:
        cursor.execute("""
            SELECT TOP 1
                userID,
                username,
                full_name,
                email,
                role,
                user_password
            FROM user_account
            WHERE username = ?
        """, (username,))
        row = cursor.fetchone()
        if not row:
            return None
        cols = [c[0] for c in cursor.description]
        return dict(zip(cols, row))
    except Exception as e:
        print("[DB ERROR]", e)
        return None
    finally:
        cursor.close()
        connection.close()