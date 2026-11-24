import pytest


@pytest.fixture(autouse=True)
def disable_update_checks(monkeypatch):
    monkeypatch.setenv("OUTLOOK_CLI_NO_UPDATE_CHECK", "1")
