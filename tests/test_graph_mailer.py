"""Tests für den externen Microsoft-Graph-Mailtransport."""
import json
import socket
import threading

from app import graph_mailer, mailer


def test_graph_socket_protocol_and_attachments(tmp_path):
    path = str(tmp_path / "mail.sock")
    received = {}
    ready = threading.Event()

    def server():
        with socket.socket(socket.AF_UNIX, socket.SOCK_STREAM) as listener:
            listener.bind(path)
            listener.listen(1)
            ready.set()
            conn, _ = listener.accept()
            with conn:
                data = bytearray()
                while b"\n" not in data:
                    chunk = conn.recv(4096)
                    if not chunk:
                        break
                    data.extend(chunk)
                received.update(json.loads(bytes(data).split(b"\n", 1)[0].decode("utf-8")))
                conn.sendall(b'{"ok":true}\n')

    thread = threading.Thread(target=server, daemon=True)
    thread.start()
    assert ready.wait(timeout=2)
    graph_mailer.send(
        "test@example.de", "Test", "Nachricht",
        [("bericht.pdf", b"%PDF-test", "application/pdf")],
        socket_path=path,
    )
    thread.join(timeout=2)
    assert received["to"] == "test@example.de"
    assert received["subject"] == "Test"
    assert received["body"] == "Nachricht"
    assert received["attachments"] == [{
        "name": "bericht.pdf",
        "mime": "application/pdf",
        "content": "JVBERi10ZXN0",
    }]


def test_graph_socket_error_response(tmp_path):
    path = str(tmp_path / "mail.sock")

    ready = threading.Event()

    def server():
        with socket.socket(socket.AF_UNIX, socket.SOCK_STREAM) as listener:
            listener.bind(path)
            listener.listen(1)
            ready.set()
            conn, _ = listener.accept()
            with conn:
                while conn.recv(4096):
                    pass
                conn.sendall(b'{"ok":false,"error":"Graph API 403"}\n')

    thread = threading.Thread(target=server, daemon=True)
    thread.start()
    assert ready.wait(timeout=2)
    try:
        graph_mailer.send("test@example.de", "Test", "Text", socket_path=path)
        assert False, "GraphMailError erwartet"
    except graph_mailer.GraphMailError as exc:
        assert "Graph API 403" in str(exc)
    thread.join(timeout=2)


def test_graph_mail_transport_dispatch(monkeypatch):
    called = {}

    def fake_send(to, subject, body, attachments, socket_path):
        called.update(to=to, subject=subject, body=body, socket_path=socket_path)

    monkeypatch.setattr(mailer, "send_graph_mail", fake_send)
    monkeypatch.setattr(mailer.settings, "mail_transport", "graph_socket")
    monkeypatch.setattr(mailer.settings, "graph_socket", "/run/graphmailer/mail.sock")
    mailer.send_mail(None, "test@example.de", "Betreff", "Text")
    assert called == {
        "to": "test@example.de",
        "subject": "Betreff",
        "body": "Text",
        "socket_path": "/run/graphmailer/mail.sock",
    }
