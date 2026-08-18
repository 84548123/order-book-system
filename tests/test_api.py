from pathlib import Path

from fastapi.testclient import TestClient

from app.main import app


client = TestClient(app)


def test_health():
    assert client.get("/health").json() == {"status": "healthy"}


def test_rejects_wrong_file_type():
    response = client.post("/api/process", files={"file": ("orders.csv", b"a,b", "text/csv")})
    assert response.status_code == 400


def test_latest_returns_404_without_saved_data(monkeypatch):
    monkeypatch.setattr("app.main.LATEST_DATA_PATH", Path("tests/fixtures/missing.json"))
    response = client.get("/api/latest")
    assert response.status_code == 404


def test_latest_returns_shared_data(monkeypatch):
    monkeypatch.setattr("app.main.LATEST_DATA_PATH", Path("tests/fixtures/latest.json"))
    response = client.get("/api/latest")
    assert response.status_code == 200
    assert response.json()["records"][0]["PoNo"] == "123"
