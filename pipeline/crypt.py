"""Files kept in the public repository are encrypted with the staff password: b'ITC1' + salt(16) + iv(12) + AES-GCM ciphertext, key = PBKDF2-SHA256(password, salt, 200000)."""
import os
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
MAGIC = b'ITC1'; ITER = 200000

def _key(pw, salt):
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pw.encode())

def enc(data: bytes, pw: str) -> bytes:
    salt, iv = os.urandom(16), os.urandom(12)
    return MAGIC + salt + iv + AESGCM(_key(pw, salt)).encrypt(iv, data, None)

def dec(blob: bytes, pw: str) -> bytes:
    assert blob[:4] == MAGIC, 'not an ITC1 file'
    salt, iv = blob[4:20], blob[20:32]
    return AESGCM(_key(pw, salt)).decrypt(iv, blob[32:], None)

def read_enc(path, pw):
    """A file may be stored whole (x.enc) or in parts (x.enc.part0, x.enc.part1, ...)."""
    if os.path.exists(path): return dec(open(path, 'rb').read(), pw)
    d, b = os.path.split(path); parts = sorted((f for f in os.listdir(d or '.') if f.startswith(b + '.part')), key=lambda f: int(f.rsplit('part', 1)[1]))
    if not parts: return None
    return dec(b''.join(open(os.path.join(d, f), 'rb').read() for f in parts), pw)
