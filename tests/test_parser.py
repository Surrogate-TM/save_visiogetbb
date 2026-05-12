import os
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from parser import ForumParser, detect_extension_from_response, normalize_url, url_to_local_path


class DummyResponse:
    def __init__(self, headers):
        self.headers = headers


def test_normalize_url_strips_session_and_sorting_params():
    assert (
        normalize_url("https://visio.getbb.ru/viewtopic.php?t=1&sid=abc&start=20&st=0")
        == "https://visio.getbb.ru/viewtopic.php?t=1&start=20"
    )


def test_url_to_local_path_preserves_query_in_filename(tmp_path):
    path = url_to_local_path(
        "https://visio.getbb.ru/viewtopic.php?f=3&t=1571&start=40",
        tmp_path,
    )

    assert path == tmp_path / "viewtopic__f=3&t=1571&start=40.html"


def test_detect_extension_prefers_content_disposition_filename():
    response = DummyResponse(
        {"Content-Disposition": 'attachment; filename="diagram.vsd"', "Content-Type": "application/octet-stream"}
    )

    assert detect_extension_from_response(response, "https://visio.getbb.ru/download/file.php?id=7") == ".vsd"


def test_process_page_rewrites_forum_links_and_downloads_assets(tmp_path):
    parser = ForumParser(str(tmp_path), delay=0)
    downloaded = {
        "https://visio.getbb.ru/download/file.php?id=7": tmp_path / "download" / "file_7.7z",
        "https://example.com/image.png": tmp_path / "images_cache" / "example_com" / "image.png",
    }

    def fake_download_file(url):
        return downloaded[url]

    def fake_download_asset(url):
        return downloaded[url]

    parser.download_file = fake_download_file
    parser.download_asset = fake_download_asset

    html = """
    <html><body>
      <a href="viewforum.php?f=3&amp;sid=abc">Forum</a>
      <a href="memberlist.php">Members</a>
      <a href="download/file.php?id=7">Attachment</a>
      <img src="https://example.com/image.png">
    </body></html>
    """

    processed = parser.process_page("https://visio.getbb.ru/viewtopic.php?t=1", html)

    assert 'href="viewforum__f=3.html"' in processed
    assert 'href="#"' in processed
    assert 'href="download/file_7.7z"' in processed
    assert 'src="images_cache/example_com/image.png"' in processed
    assert list(parser.queue) == ["https://visio.getbb.ru/viewforum.php?f=3"]
