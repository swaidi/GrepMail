import hashlib, io, os, sys, tempfile, urllib.error
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import gm_update as u


def release(tag="v2.1.0", assets=None, body="Fixes and improvements"):
    return {"tag_name": tag, "body": body, "html_url": "https://github.com/x/y/releases/tag/" + tag,
            "assets": assets if assets is not None else []}


def asset(name, data=b"payload"):
    return {"name": name, "size": len(data),
            "browser_download_url": "https://github.com/x/y/releases/download/v/" + name,
            "digest": "sha256:" + hashlib.sha256(data).hexdigest()}


def test_versions():
    assert u.parse_version("v2.10.1") == (2, 10, 1)
    assert u.is_newer("v2.1.0", "2.0.0") and u.is_newer("2.0.1", "2.0.0") and u.is_newer("3", "2.9.9")
    assert not u.is_newer("2.0.0", "2.0.0") and not u.is_newer("v1.9.9", "2.0.0")
    assert u.is_newer("2.10.0", "2.9.0")
    print("versions ok")


def test_check_latest():
    assets = [asset("GrepMail-Setup-2.1.0.exe"), asset("grepmail_2.1.0_amd64.deb")]
    r = u.check_latest("2.0.0", fetch=lambda url: release(assets=assets), windows=True)
    assert r["version"] == "2.1.0" and r["asset"]["name"].endswith(".exe") and r["notes"]
    assert len(r["asset"]["sha256"]) == 64
    r = u.check_latest("2.0.0", fetch=lambda url: release(assets=assets), windows=False)
    assert r["asset"]["name"].endswith(".deb")
    assert u.check_latest("2.1.0", fetch=lambda url: release(assets=assets)) is None
    assert u.check_latest("2.0.0", fetch=lambda url: release(assets=[]), windows=True)["asset"] is None
    evil = [{"name": "GrepMail-Setup.exe", "size": 1, "browser_download_url": "https://evil.example/x.exe"}]
    assert u.check_latest("2.0.0", fetch=lambda url: release(assets=evil), windows=True)["asset"] is None

    def nf(url):
        raise urllib.error.HTTPError(url, 404, "nf", {}, None)
    assert u.check_latest("2.0.0", fetch=nf) is None
    print("check latest ok")


class FakeResp(io.BytesIO):
    def __init__(self, data, url):
        super().__init__(data)
        self.headers = {"Content-Length": str(len(data))}
        self._url = url
    def geturl(self): return self._url
    def __enter__(self): return self
    def __exit__(self, *a): return False


def test_download():
    data = b"x" * 600000
    a = u.check_latest("2.0.0", fetch=lambda url: release(assets=[asset("GrepMail-Setup-2.1.0.exe", data)]),
                       windows=True)["asset"]
    d = tempfile.mkdtemp()
    seen = []
    p = u.download(a, d, progress=lambda x, y: seen.append((x, y)),
                   opener=lambda req: FakeResp(data, a["url"]))
    assert open(p, "rb").read() == data and seen[-1][0] == len(data)
    bad = dict(a, sha256="0" * 64)
    try:
        u.download(bad, d, opener=lambda req: FakeResp(data, a["url"])); raise SystemExit("must reject checksum")
    except ValueError:
        assert not os.path.exists(p)
    try:
        u.download(a, d, opener=lambda req: FakeResp(data, "https://evil.example/x")); raise SystemExit("must reject redirect")
    except ValueError:
        pass
    try:
        u.download(dict(a, url="http://github.com/x"), d); raise SystemExit("must reject http")
    except ValueError:
        pass
    print("download ok")


if __name__ == "__main__":
    test_versions(); test_check_latest(); test_download()
    print("ALL UPDATE TESTS PASSED")
