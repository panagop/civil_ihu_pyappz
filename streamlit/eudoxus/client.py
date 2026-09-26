"""Minimal client for the Eudoxus (service.eudoxus.gr) book-search backend.

This is an UNOFFICIAL, reverse-engineered client: it calls the same JSON
endpoint that the public «Σύνθετη Αναζήτηση» page at
https://service.eudoxus.gr/search/#/advanced calls from the browser. There is
no published contract, so treat every response as best-effort and never let a
failure here take down a page — :func:`fetch_books` reports errors per book
rather than raising.

Endpoint
--------
    PUT https://service.eudoxus.gr/search/rest/app/advanced-search
    Content-Type: application/json

The body is a ``SearchBooksModel``. Unknown keys are rejected with HTTP 500,
so only send fields the server knows. ``first`` and ``pageSize`` are required.

Response
--------
    {"start": 0, "total": <int>, "results": [ {...book...} ], "facets": null}

Grew out of ``files/eudoxus/eudoxus.py``, which stays as the standalone script
it was written as.
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Any, Callable, Iterator

import requests

BASE = "https://service.eudoxus.gr/search/rest/app"
ENDPOINT = f"{BASE}/advanced-search"

# Fields the page itself sends. Sending exactly these is the safest baseline:
# anything the server does not recognise comes back as HTTP 500.
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

# What we keep out of the ~70 fields a book carries. The rest is pricing,
# workflow and file-path noise that means nothing outside Eudoxus.
BOOK_FIELDS = {
    "id": "book_id",
    "title": "title",
    "subtitle": "subtitle",
    "authors": "authors",
    "isbn": "isbn",
    "publisherName": "publisher",
    "publicationYear": "publication_year",
    "editionNumber": "edition_number",
    "active": "active",
    "selectable": "selectable",
}


class EudoxusError(RuntimeError):
    pass


@dataclass
class Eudoxus:
    delay: float = 0.15  # be polite; this is a public service
    timeout: float = 30.0
    session: requests.Session = field(default=None, repr=False)

    def __post_init__(self) -> None:
        self.session = self.session or requests.Session()
        self.session.headers.update(
            {
                "Content-Type": "application/json",
                "Accept": "application/json",
                # Identify yourself. Don't pretend to be a browser.
                "User-Agent": "civil-ihu-pyappz/1.0 (course reading list audit)",
            }
        )

    # -- low level ---------------------------------------------------------
    def search(self, *, first: int = 0, page_size: int = 9, **filters: Any) -> dict:
        """One raw call. ``filters`` are SearchBooksModel fields."""
        body = {**_TEMPLATE, **filters, "first": first, "pageSize": page_size}
        response = self.session.put(ENDPOINT, json=body, timeout=self.timeout)
        if response.status_code != 200:
            raise EudoxusError(f"HTTP {response.status_code}: {response.text[:300]}")
        if self.delay:
            time.sleep(self.delay)
        return response.json()

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
        """Look up one book by its Eudoxus code. ``id`` is an EXACT match."""
        page = self.search(id=str(book_id), page_size=1)
        results = page.get("results") or []
        return results[0] if results else None

    def fetch_books(
        self,
        book_ids: list[int],
        progress: Callable[[int, int], None] | None = None,
    ) -> list[dict]:
        """Catalogue rows for many books, one request each.

        There is no bulk endpoint — ``id`` is an exact match on a single code —
        so this is inherently sequential and costs roughly a second per book.
        That is why the result is stored rather than fetched on every page view.

        Never raises: a book that errors or does not exist comes back with
        ``found=False`` and the reason, so one dead code cannot abort a check of
        two hundred.
        """
        rows: list[dict] = []
        total = len(book_ids)
        for index, book_id in enumerate(book_ids, 1):
            row: dict[str, Any] = {"book_id": int(book_id), "found": False, "error": None}
            try:
                book = self.get_book(book_id)
            except (EudoxusError, requests.RequestException) as exc:
                row["error"] = str(exc)[:200]
            else:
                if book is None:
                    row["error"] = "δεν βρέθηκε στο μητρώο"
                else:
                    row.update(
                        {target: book.get(source) for source, target in BOOK_FIELDS.items()}
                    )
                    row["book_id"] = int(book.get("id", book_id))
                    row["found"] = True
            rows.append(row)
            if progress:
                progress(index, total)
        return rows

    def search_books(self, *, page_size: int = 25, **filters: Any) -> list[dict]:
        """Catalogue rows for a title/author search, for the 'add a book' UI."""
        page = self.search(page_size=page_size, **filters)
        return [
            {
                **{target: book.get(source) for source, target in BOOK_FIELDS.items()},
                "found": True,
                "error": None,
            }
            for book in page.get("results") or []
        ]

    def is_available(self, book_id: str | int) -> tuple[bool, str]:
        """(ok, reason) for a single book code."""
        book = self.get_book(book_id)
        if book is None:
            return False, "δεν βρέθηκε στο μητρώο"
        if not book.get("active", False):
            return False, "ανενεργό (active=false)"
        if not book.get("selectable", False):
            return False, "μη επιλέξιμο (selectable=false)"
        return True, "ενεργό και επιλέξιμο"


def availability_reason(row: dict) -> str:
    """Why a stored catalogue row is or is not usable next year."""
    if not row.get("found", True):
        return row.get("error") or "δεν βρέθηκε στο μητρώο"
    if not row.get("active"):
        return "ανενεργό (active=false)"
    if not row.get("selectable"):
        return "μη επιλέξιμο (selectable=false)"
    return ""
