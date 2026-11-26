import httpx


def create_sharepoint_client(bearer_token: str) -> httpx.AsyncClient:
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {bearer_token}"
    }
    return httpx.AsyncClient(
        base_url="https://graph.microsoft.com/v1.0/",
        headers=headers,
        timeout=httpx.Timeout(20),
        limits=httpx.Limits(max_connections=10, max_keepalive_connections=5)
    )
