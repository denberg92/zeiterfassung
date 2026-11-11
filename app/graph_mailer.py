"""Client für den externen Microsoft-Graph-Maildienst über einen Unix-Socket.

Das Zertifikat und die Microsoft-Graph-Authentifizierung bleiben bewusst außerhalb
des SPK-Pakets. Die App übergibt nur Empfänger, Text und optionale Anhänge.
"""
import base64
import json
import os
import socket
from typing import List, Optional, Tuple


MAX_REQUEST = 12 * 1024 * 1024
MAX_ATTACHMENT = 3 * 1024 * 1024
MAX_RESPONSE = 16 * 1024
SOCKET_TIMEOUT = 90


class GraphMailError(Exception):
    """Fehler bei Übergabe oder Versand über den Graph-Maildienst."""


def send(to: str, subject: str, body: str,
         attachments: Optional[List[Tuple[str, bytes, str]]] = None,
         socket_path: str = "/run/graphmailer/mail.sock") -> None:
    """Übergibt eine Mail synchron an den lokalen Graph-Maildienst."""
    if not socket_path or not os.path.isabs(socket_path):
        raise GraphMailError("Der Graph-Mail-Socket muss als absoluter Pfad angegeben sein")

    encoded_attachments = []
    for name, content, mime in attachments or []:
        if not isinstance(content, bytes):
            raise GraphMailError(f"Anhang {name!r} liegt nicht als Bytes vor")
        if len(content) > MAX_ATTACHMENT:
            raise GraphMailError(f"Anhang {name!r} ist größer als 3 MiB")
        if not name or not mime or "/" not in mime:
            raise GraphMailError("Name oder MIME-Typ eines Anhangs ist ungültig")
        encoded_attachments.append({
            "name": name,
            "mime": mime,
            "content": base64.b64encode(content).decode("ascii"),
        })

    payload = {
        "to": to,
        "subject": subject,
        "body": body,
        "attachments": encoded_attachments,
    }
    try:
        request = (json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "\n").encode("utf-8")
    except (TypeError, ValueError) as e:
        raise GraphMailError(f"Maildaten konnten nicht serialisiert werden: {e}")
    if len(request) > MAX_REQUEST:
        raise GraphMailError("Mailauftrag ist größer als 12 MiB")

    try:
        with socket.socket(socket.AF_UNIX, socket.SOCK_STREAM) as sock:
            sock.settimeout(SOCKET_TIMEOUT)
            sock.connect(socket_path)
            sock.sendall(request)
            sock.shutdown(socket.SHUT_WR)

            response = bytearray()
            while len(response) <= MAX_RESPONSE:
                chunk = sock.recv(4096)
                if not chunk:
                    break
                response.extend(chunk)
                if b"\n" in response:
                    response = response.split(b"\n", 1)[0]
                    break
            if not response:
                raise GraphMailError("Der Graph-Maildienst hat keine Antwort geliefert")
            if len(response) > MAX_RESPONSE:
                raise GraphMailError("Antwort des Graph-Maildienstes ist zu groß")
            try:
                result = json.loads(response.decode("utf-8"))
            except (UnicodeDecodeError, ValueError) as e:
                raise GraphMailError(f"Ungültige Antwort des Graph-Maildienstes: {e}")
            if not isinstance(result, dict) or result.get("ok") is not True:
                detail = result.get("error", "Unbekannter Fehler") if isinstance(result, dict) else "Ungültiges Antwortformat"
                raise GraphMailError(f"Graph-Mailversand fehlgeschlagen: {str(detail)[:500]}")
    except GraphMailError:
        raise
    except socket.timeout:
        raise GraphMailError("Zeitüberschreitung beim Warten auf den Graph-Maildienst")
    except OSError as e:
        raise GraphMailError(f"Graph-Maildienst nicht erreichbar ({e})")
