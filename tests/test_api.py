from fastapi.testclient import TestClient

from app.main import app


client = TestClient(app)


def test_health():
    assert client.get("/health").json() == {"status": "healthy"}


def test_rejects_wrong_file_type():
    response = client.post("/api/process", files={"file": ("orders.csv", b"a,b", "text/csv")})
    assert response.status_code == 400

