"""In-app update: version logic, URL allowlist, checksum verification, routes (network simulated)."""
from __future__ import annotations

import hashlib
import io
import json
import time

import pytest

from personnel_evaluation_report import VERSION, updater


def test_version_compare():
    assert updater.is_newer("0.2.2", "0.2.1") and updater.is_newer("v1.0", "0.9.9")
    assert not updater.is_newer("0.2.1", "0.2.1") and not updater.is_newer("0.2", "0.2.0")
    with pytest.raises(ValueError):
        updater.parse_version("latest")


@pytest.mark.parametrize("url", ["http://github.com/x", "https://evil.example/x", "https://github.com.evil.io/x",
                                 "file:///etc/passwd"])
def test_only_github_https(url):
    with pytest.raises(updater.UpdateError):
        updater._check_url(url)


def test_allowed_urls():
    updater._check_url("https://api.github.com/repos/a/b/releases/latest")
    updater._check_url("https://release-assets.githubusercontent.com/x")


def test_expected_sha256():
    h = "a" * 64
    assert updater.expected_sha256(f"{h}  PersonnelEvaluationReport-Setup-9.9.9.exe\n", "PersonnelEvaluationReport-Setup-9.9.9.exe") == h
    with pytest.raises(updater.UpdateError):
        updater.expected_sha256(f"{h}  other.exe\n", "PersonnelEvaluationReport-Setup-9.9.9.exe")


class FakeResp(io.BytesIO):
    def __init__(self, data: bytes, url: str):
        super().__init__(data)
        self._url = url
        self.headers = {"Content-Length": str(len(data))}

    def geturl(self):
        return self._url

    def __enter__(self):
        return self

    def __exit__(self, *a):
        self.close()


def fake_release(monkeypatch, tmp_path, payload: bytes, good_sum=True, version="9.9.9"):
    name = updater.asset_name(version) or f"PersonnelEvaluationReport-Setup-{version}.exe"
    monkeypatch.setattr(updater, "asset_name", lambda v: name)
    monkeypatch.setattr(updater, "user_data_dir", lambda: tmp_path)
    base = f"https://github.com/irahog/personnel-evaluation-report-releases/releases/download/v{version}/"
    sha = hashlib.sha256(payload if good_sum else b"other").hexdigest()
    files = {
        updater.API_LATEST: json.dumps({"tag_name": f"v{version}", "html_url": updater.RELEASES_PAGE, "assets": [
            {"name": name, "browser_download_url": base + name},
            {"name": "SHA256SUMS.txt", "browser_download_url": base + "SHA256SUMS.txt"}]}).encode(),
        base + name: payload,
        base + "SHA256SUMS.txt": f"{sha}  {name}\n".encode(),
    }

    def opener(url):
        updater._check_url(url)
        if url not in files:
            raise OSError("404")
        return FakeResp(files[url], url)
    monkeypatch.setattr(updater, "_open", opener)
    return name


def test_check_and_download_verifies(monkeypatch, tmp_path):
    name = fake_release(monkeypatch, tmp_path, b"installer-bytes" * 1000)
    info = updater.check()
    assert info.newer and info.version == "9.9.9" and info.asset == name
    job = updater.Job(state="downloading")
    updater.download(info, job)
    assert job.state == "ready" and job.file.read_bytes() == b"installer-bytes" * 1000
    assert job.file.parent == tmp_path / "updates"


def test_download_rejects_bad_checksum(monkeypatch, tmp_path):
    fake_release(monkeypatch, tmp_path, b"tampered", good_sum=False)
    job = updater.Job(state="downloading")
    updater.download(updater.check(), job)
    assert job.state == "error" and "SHA-256" in job.message
    assert not list((tmp_path / "updates").glob("*"))


def test_check_offline(monkeypatch):
    def boom(url):
        raise OSError("no network")
    monkeypatch.setattr(updater, "_open", boom)
    with pytest.raises(updater.UpdateError, match="اینترنت"):
        updater.check()


def test_same_version_is_not_newer(monkeypatch, tmp_path):
    fake_release(monkeypatch, tmp_path, b"x", version=VERSION)
    assert not updater.check().newer


def test_update_routes(settings, monkeypatch, tmp_path):
    from personnel_evaluation_report.webapp import create_app
    from tests.conftest import Client
    fake_release(monkeypatch, tmp_path, b"payload")
    monkeypatch.setattr(updater, "can_self_update", lambda: True)
    launched, stopped = [], []
    monkeypatch.setattr(updater, "prepare_install", lambda f: ["installer", str(f)])
    monkeypatch.setattr(updater, "launch", lambda cmd: launched.append(cmd))
    app = create_app(settings, shutdown=lambda: stopped.append(1))
    c = Client(app, [])
    body = c.post("/settings/update/check", follow_redirects=True).get_data(as_text=True)
    assert "نسخهٔ تازهٔ" in body and "دانلود نسخهٔ" in body
    c.post("/settings/update/download")
    for _ in range(50):
        if c.get("/settings/update/status").json["state"] == "ready":
            break
        time.sleep(0.05)
    assert c.get("/settings/update/status").json["percent"] == 100
    assert "نصب و راه‌اندازی دوباره" in c.get("/settings").get_data(as_text=True)
    r = c.post("/settings/update/install")
    assert "در حال نصب نسخهٔ" in r.get_data(as_text=True) and launched
    time.sleep(1.2)
    assert stopped and app.extensions["personnel_evaluation_report"].sessions == {}


def test_settings_page_shows_update_card(client):
    page = client.get("/settings").get_data(as_text=True)
    assert 'id="update"' in page and "بررسی نسخهٔ جدید" in page
