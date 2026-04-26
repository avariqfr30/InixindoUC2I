import hashlib
import logging
import os
import secrets
import sqlite3
from datetime import datetime, timezone, timedelta

from flask import g, request, session
from werkzeug.security import check_password_hash, generate_password_hash

from config import (
    APP_SECRET_KEY,
    AUTH_DB_PATH,
    SESSION_ACTIVITY_TOUCH_SECONDS,
    SESSION_IDLE_TIMEOUT_SECONDS,
    SESSION_MAX_ACTIVE_PER_USER,
    SESSION_MAX_ACTIVE_TOTAL,
)

logger = logging.getLogger(__name__)
PASSWORD_HASH_METHOD = "pbkdf2:sha256"
SESSION_ID_BYTES = 32


class ActiveSessionCapacityError(RuntimeError):
    pass


def _auth_connection():
    connection = sqlite3.connect(AUTH_DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def _utc_now():
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _utc_in(seconds):
    return (
        datetime.now(timezone.utc).replace(microsecond=0) + timedelta(seconds=seconds)
    ).isoformat().replace("+00:00", "Z")


def _parse_utc_timestamp(raw_value):
    value = str(raw_value or "").strip()
    if not value:
        return None

    normalized = value.replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def _hash_fingerprint(raw_value):
    clean = str(raw_value or "").strip()
    if not clean:
        return None
    digest = hashlib.sha256(f"{APP_SECRET_KEY}|{clean}".encode("utf-8")).hexdigest()
    return digest[:24]


def _request_ip():
    forwarded = (request.headers.get("X-Forwarded-For") or "").split(",")[0].strip()
    return forwarded or (request.remote_addr or "").strip()


def _cleanup_stale_sessions(connection, now_iso=None):
    reference_time = now_iso or _utc_now()
    connection.execute(
        """
        DELETE FROM user_sessions
        WHERE revoked_at IS NOT NULL OR expires_at <= ?
        """,
        (reference_time,),
    )


def _count_active_sessions(connection, now_iso, user_id=None):
    if user_id is None:
        row = connection.execute(
            """
            SELECT COUNT(*) AS total
            FROM user_sessions
            WHERE revoked_at IS NULL AND expires_at > ?
            """,
            (now_iso,),
        ).fetchone()
    else:
        row = connection.execute(
            """
            SELECT COUNT(*) AS total
            FROM user_sessions
            WHERE user_id = ? AND revoked_at IS NULL AND expires_at > ?
            """,
            (user_id, now_iso),
        ).fetchone()
    return int(row["total"]) if row else 0


def _session_stats_for_user(user_id):
    now_iso = _utc_now()
    with _auth_connection() as connection:
        _cleanup_stale_sessions(connection, now_iso)
        total_active = _count_active_sessions(connection, now_iso)
        user_active = _count_active_sessions(connection, now_iso, user_id=user_id)
        connection.commit()
        return {
            "total_active": total_active,
            "user_active": user_active,
            "max_total": SESSION_MAX_ACTIVE_TOTAL,
            "max_per_user": SESSION_MAX_ACTIVE_PER_USER,
        }


def _session_capacity_snapshot():
    now_iso = _utc_now()
    with _auth_connection() as connection:
        _cleanup_stale_sessions(connection, now_iso)
        total_active = _count_active_sessions(connection, now_iso)
        connection.commit()
    return {
        "total_active": total_active,
        "max_total": SESSION_MAX_ACTIVE_TOTAL,
        "max_per_user": SESSION_MAX_ACTIVE_PER_USER,
        "idle_timeout_seconds": SESSION_IDLE_TIMEOUT_SECONDS,
    }


def _session_stats_for_username(username):
    user = _get_user_by_username(username)
    if not user:
        return {
            "total_active": 0,
            "user_active": 0,
            "max_total": SESSION_MAX_ACTIVE_TOTAL,
            "max_per_user": SESSION_MAX_ACTIVE_PER_USER,
        }
    return _session_stats_for_user(int(user["id"]))


def _issue_session(user):
    now_iso = _utc_now()
    expires_at = _utc_in(SESSION_IDLE_TIMEOUT_SECONDS)
    new_session_id = secrets.token_urlsafe(SESSION_ID_BYTES)
    user_id = int(user["id"])
    ip_hash = _hash_fingerprint(_request_ip())
    user_agent_hash = _hash_fingerprint(request.headers.get("User-Agent"))
    revoked_count = 0

    with _auth_connection() as connection:
        _cleanup_stale_sessions(connection, now_iso)

        active_user_sessions = connection.execute(
            """
            SELECT session_id
            FROM user_sessions
            WHERE user_id = ? AND revoked_at IS NULL AND expires_at > ?
            ORDER BY last_seen_at ASC, created_at ASC
            """,
            (user_id, now_iso),
        ).fetchall()

        if SESSION_MAX_ACTIVE_PER_USER > 0:
            overflow = len(active_user_sessions) - SESSION_MAX_ACTIVE_PER_USER + 1
            if overflow > 0:
                sessions_to_remove = [
                    (row["session_id"],) for row in active_user_sessions[:overflow]
                ]
                revoked_count = overflow
                connection.executemany(
                    "DELETE FROM user_sessions WHERE session_id = ?",
                    sessions_to_remove,
                )

        if SESSION_MAX_ACTIVE_TOTAL > 0:
            active_total = _count_active_sessions(connection, now_iso)
            if active_total >= SESSION_MAX_ACTIVE_TOTAL:
                raise ActiveSessionCapacityError(
                    "Batas sesi aktif sudah tercapai. Coba lagi beberapa saat lagi atau logout dari perangkat lain."
                )

        connection.execute(
            """
            INSERT INTO user_sessions (
                session_id,
                user_id,
                username,
                created_at,
                last_seen_at,
                expires_at,
                revoked_at,
                revocation_reason,
                ip_hash,
                user_agent_hash
            )
            VALUES (?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?)
            """,
            (
                new_session_id,
                user_id,
                user["username"],
                now_iso,
                now_iso,
                expires_at,
                ip_hash,
                user_agent_hash,
            ),
        )
        connection.commit()

    return new_session_id, revoked_count


def _revoke_session(session_id, reason="logout"):
    clean_session_id = str(session_id or "").strip()
    if not clean_session_id:
        return
    with _auth_connection() as connection:
        connection.execute(
            """
            UPDATE user_sessions
            SET revoked_at = ?, revocation_reason = ?
            WHERE session_id = ?
            """,
            (_utc_now(), reason, clean_session_id),
        )
        connection.commit()


def _touch_session(username, session_id):
    clean_username = str(username or "").strip()
    clean_session_id = str(session_id or "").strip()
    if not clean_username or not clean_session_id:
        return None

    now_iso = _utc_now()
    with _auth_connection() as connection:
        _cleanup_stale_sessions(connection, now_iso)
        active_session = connection.execute(
            """
            SELECT s.user_id, s.username, s.last_seen_at
            FROM user_sessions AS s
            INNER JOIN users AS u ON u.id = s.user_id
            WHERE s.username = ? AND s.session_id = ? AND s.revoked_at IS NULL AND s.expires_at > ?
            LIMIT 1
            """,
            (clean_username, clean_session_id, now_iso),
        ).fetchone()
        if not active_session:
            connection.commit()
            return None

        should_touch = True
        if SESSION_ACTIVITY_TOUCH_SECONDS > 0:
            last_seen = _parse_utc_timestamp(active_session["last_seen_at"])
            if last_seen is not None:
                elapsed = (datetime.now(timezone.utc) - last_seen).total_seconds()
                should_touch = elapsed >= SESSION_ACTIVITY_TOUCH_SECONDS

        if should_touch:
            connection.execute(
                """
                UPDATE user_sessions
                SET last_seen_at = ?, expires_at = ?
                WHERE session_id = ?
                """,
                (now_iso, _utc_in(SESSION_IDLE_TIMEOUT_SECONDS), clean_session_id),
            )
        connection.commit()
        return active_session["username"]


def _start_authenticated_session(user):
    session_id, revoked_count = _issue_session(user)
    session.clear()
    session.permanent = True
    session["username"] = user["username"]
    session["sid"] = session_id
    return revoked_count


def _init_auth_db():
    os.makedirs(os.path.dirname(AUTH_DB_PATH), exist_ok=True)
    with _auth_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS user_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL UNIQUE,
                user_id INTEGER NOT NULL,
                username TEXT NOT NULL,
                created_at TIMESTAMP NOT NULL,
                last_seen_at TIMESTAMP NOT NULL,
                expires_at TIMESTAMP NOT NULL,
                revoked_at TIMESTAMP,
                revocation_reason TEXT,
                ip_hash TEXT,
                user_agent_hash TEXT,
                FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
            )
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_user_sessions_active_lookup
            ON user_sessions (username, session_id, expires_at, revoked_at)
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_user_sessions_expiry
            ON user_sessions (expires_at, revoked_at)
            """
        )
        _cleanup_stale_sessions(connection)
        connection.commit()


def _get_user_by_username(username):
    clean_username = str(username or "").strip()
    if not clean_username:
        return None

    with _auth_connection() as connection:
        return connection.execute(
            "SELECT id, username, password_hash FROM users WHERE username = ?",
            (clean_username,),
        ).fetchone()


def _create_user(username, password):
    clean_username = username.strip()
    with _auth_connection() as connection:
        connection.execute(
            "INSERT INTO users (username, password_hash) VALUES (?, ?)",
            (
                clean_username,
                generate_password_hash(password, method=PASSWORD_HASH_METHOD),
            ),
        )
        connection.commit()
    return _get_user_by_username(clean_username)


def _user_count():
    with _auth_connection() as connection:
        row = connection.execute("SELECT COUNT(*) AS total FROM users").fetchone()
        return int(row["total"]) if row else 0


def _verify_password(password_hash, password):
    try:
        return check_password_hash(password_hash, password)
    except AttributeError:
        if str(password_hash or "").startswith("scrypt:"):
            logger.warning(
                "Stored password hash uses scrypt, but this Python build does not support hashlib.scrypt."
            )
            return False
        raise


def _current_user():
    if hasattr(g, "_current_user_cached"):
        return g._current_user_cached

    username = session.get("username")
    session_id = session.get("sid")
    if isinstance(username, str) and username.strip() and isinstance(session_id, str):
        active_username = _touch_session(username, session_id)
        if active_username:
            g._current_user_cached = active_username
            return active_username

    session.clear()
    g._current_user_cached = None
    return None


def logout_current_session(reason="manual_logout"):
    _revoke_session(session.get("sid"), reason=reason)
    session.clear()


init_auth_db = _init_auth_db
current_user = _current_user
get_user_by_username = _get_user_by_username
create_user = _create_user
user_count = _user_count
verify_password = _verify_password
start_authenticated_session = _start_authenticated_session
revoke_session = _revoke_session
session_capacity_snapshot = _session_capacity_snapshot
session_stats_for_username = _session_stats_for_username
