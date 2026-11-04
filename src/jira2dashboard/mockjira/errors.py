"""HTTP error carried from the API functions to the server."""


class ApiError(Exception):
    """An error response: status, JSON body (or plain text) and extra headers (for example Retry-After)."""

    def __init__(self, status: int, body, headers: dict[str, str] | None = None):
        super().__init__(f"{status} {body}")
        self.status = status
        self.body = body
        self.headers = headers or {}
