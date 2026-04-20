"""
auth.py — JWT authentication for Payless Automation Hub
"""

import os
import jwt
from datetime import datetime, timedelta, timezone
from functools import wraps
from flask import request, jsonify, g

SECRET_KEY        = os.environ.get('JWT_SECRET', 'payless-hub-change-in-prod-2026')
TOKEN_EXPIRY_HOURS = 8


def create_token(user: dict) -> str:
    payload = {
        'user_id':         user['id'],
        'username':        user['username'],
        'is_admin':        bool(user['is_admin']),
        'allowed_modules': user.get('allowed_modules', ['*']),
        'exp': datetime.now(timezone.utc) + timedelta(hours=TOKEN_EXPIRY_HOURS),
    }
    return jwt.encode(payload, SECRET_KEY, algorithm='HS256')


def decode_token(token: str) -> dict:
    return jwt.decode(token, SECRET_KEY, algorithms=['HS256'])


def _get_token_from_request():
    # Try Authorization header first (regular fetch calls)
    header = request.headers.get('Authorization', '')
    if header.startswith('Bearer '):
        return header[7:]
    # Fall back to query param (EventSource can't set headers)
    return request.args.get('token') or None


def require_auth(f):
    """Decorator: endpoint requires a valid JWT. Sets g.user_id, g.username, g.is_admin, g.allowed_modules."""
    @wraps(f)
    def decorated(*args, **kwargs):
        token = _get_token_from_request()
        if not token:
            return jsonify({'error': 'Authentication required', 'code': 'no_token'}), 401
        try:
            payload     = decode_token(token)
            g.user_id   = payload['user_id']
            g.username  = payload['username']
            g.is_admin  = payload['is_admin']
            g.allowed_modules = payload.get('allowed_modules', ['*'])
        except jwt.ExpiredSignatureError:
            return jsonify({'error': 'Session expired — please log in again.', 'code': 'expired'}), 401
        except jwt.InvalidTokenError:
            return jsonify({'error': 'Invalid session token.', 'code': 'invalid'}), 401
        return f(*args, **kwargs)
    return decorated


def require_admin(f):
    """Decorator: endpoint requires a valid JWT AND is_admin=True."""
    @wraps(f)
    @require_auth
    def decorated(*args, **kwargs):
        if not g.is_admin:
            return jsonify({'error': 'Admin access required.'}), 403
        return f(*args, **kwargs)
    return decorated
