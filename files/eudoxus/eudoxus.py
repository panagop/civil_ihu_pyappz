"""
Minimal client for the Eudoxus (service.eudoxus.gr) book-search backend.

This is an UNOFFICIAL, reverse-engineered client: it calls the same JSON
endpoint that the public "Σύνθετη Αναζήτηση" page at
https://service.eudoxus.gr/search/#/advanced calls from the browser.
There is no published contract, so treat it as best-effort.

Endpoint
--------
    PUT https://service.eudoxus.gr/search/rest/app/advanced-search
    Content-Type: application/json

The body is a `SearchBooksModel`. Unknown keys are rejected with HTTP 500,
so only send fields the server knows. `first` and `pageSize` are required
(pageSize=None -> 500).

Response
--------
    {"start": 0, "total": <int>, "results": [ {...book...} ], "facets": null}
"""

from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Any, Iterator

import requests

BASE = "https://service.eudoxus.gr/search/rest/app"
ENDPOINT = f"{BASE}/advanced-search"

# Fields the page itself sends. Sending exactly these is the safest baseline.
_TEMPLATE: dict[str, Any] = {
    "id": None,
    "type": None,
    "title": None,
    "subtitle": None,
    "authors": None,
    "isbn": None,
    "volumeTitle": None,
    "volumeNumber": None,
    "editionNumber": None,
    "publicationYear": None,
    "editorialHouse": None,
    "publisherId": None,
    "finalizedByPublisher": None,
    "keywords": None,
    "first": 0,
    "pageSize": 9,
}


class EudoxusError(RuntimeError):
    pass


@dataclass
class Eudoxus:
    delay: float = 0.15          # be polite; this is a public service
    timeout: float = 30.0
    session: requests.Session | None = None

    def __post_init__(self) -> None:
        self.session = self.session or requests.Session()
        self.session.headers.update(
            {
                "Content-Type": "application/json",
                "Accept": "application/json",
                # Identify yourself. Don't pretend to be a browser.
                "User-Agent": "eudoxus-checker/1.0 (course reading list audit)",
            }
        )

    # -- low level ---------------------------------------------------------
    def search(self, *, first: int = 0, page_size: int = 9, **filters: Any) -> dict:
        """One raw call. `filters` are SearchBooksModel fields."""
        body = {**_TEMPLATE, **filters, "first": first, "pageSize": page_size}
        r = self.session.put(ENDPOINT, json=body, timeout=self.timeout)
        if r.status_code != 200:
            raise EudoxusError(f"HTTP {r.status_code}: {r.text[:300]}")
        if self.delay:
            time.sleep(self.delay)
        return r.json()

    def iter_all(self, *, page_size: int = 50, **filters: Any) -> Iterator[dict]:
        """Page through every result for a query."""
        first = 0
        while True:
            page = self.search(first=first, page_size=page_size, **filters)
            results = page.get("results") or []
            yield from results
            first += len(results)
            if not results or first >= page.get("total", 0):
                return

    # -- convenience -------------------------------------------------------
    def get_book(self, book_id: str | int) -> dict | None:
        """Look up one book by its Eudoxus code. `id` is an EXACT match."""
        page = self.search(id=str(book_id), page_size=1)
        results = page.get("results") or []
        return results[0] if results else None

    def is_available(self, book_id: str | int) -> tuple[bool, str]:
        """(ok, reason) for a single book code."""
        book = self.get_book(book_id)
        if book is None:
            return False, "not found in the registry"
        if not book.get("active", False):
            return False, "inactive (active=false)"
        if not book.get("selectable", False):
            return False, "not selectable (selectable=false)"
        return True, "active and selectable"


if __name__ == "__main__":
    eu = Eudoxus()

    # 1. single lookup
    b = eu.get_book(41959119)
    print(b["title"], "|", b["authors"], "|", b["active"], b["selectable"])

    # 2. availability check
    for code in (41959119, 94688998, 77107076, 999999999):
        print(code, eu.is_available(code))

    # 3. substring search on title (accent- and case-insensitive)
    page = eu.search(title="Εδαφομηχανικ", page_size=5)
    print("total:", page["total"])
    for r in page["results"]:
        print("  ", r["id"], r["publicationYear"], r["title"][:60])

    # 4. server-side filter that actually works: active
    print("inactive 'Στατική':", eu.search(title="Στατική", active=False, page_size=1)["total"])

    # 5. page through everything by one publisher
    n = sum(1 for _ in eu.iter_all(authors="Πρίνος", page_size=50))
    print("books by Πρίνος:", n)
