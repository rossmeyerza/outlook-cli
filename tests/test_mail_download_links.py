"""mail download-links: share links in an email body download through Graph."""
from __future__ import annotations

import argparse

from outlook_draft import runtime
from outlook_draft.commands import mail

URL = "https://tenant.sharepoint.com/:p:/r/sites/Team/Shared%20Documents/Deck.pptx?d=w1"


class Outlook:
    closed = False

    def get_message(self, message_id):
        assert message_id == "message"
        return {"Body": {"ContentType": "HTML", "Content": f'<p>See <a href="{URL}">the deck</a> and <a href="https://example.com">x</a></p>'}}

    def close(self):
        self.closed = True


class Graph:
    closed = False
    calls: list[str] = []

    def get_share_drive_item(self, url):
        self.calls.append(url)
        return {"name": "Deck.pptx"}

    def download_share_url(self, url):
        return b"pptx-bytes"

    def close(self):
        self.closed = True


def test_download_links_uses_the_runtime_graph_client(monkeypatch, tmp_path):
    graph = Graph()
    monkeypatch.setattr(runtime, "get_client", lambda **kwargs: Outlook())
    monkeypatch.setattr(runtime, "get_graph_client", lambda: graph)
    monkeypatch.setattr(mail, "_resolve_message_id", lambda client, ref: "message")
    args = argparse.Namespace(message_id="1", dir=str(tmp_path), overwrite=False, no_spinner=True, json=False)
    mail.cmd_download_links(args)
    assert (tmp_path / "Deck.pptx").read_bytes() == b"pptx-bytes"
    assert graph.calls == [URL]
    assert graph.closed is True
