"""Utility functions for QuickBooks OAuth token refresh."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Dict

from intuitlib.client import AuthClient
from quickbooks import QuickBooks


# Determine the location of ``qb_tokens.json`` relative to this module so that
# token refresh works regardless of the current working directory used to run
# the scripts.
TOKEN_FILE = Path(__file__).with_name("qb_tokens.json")


def refresh_access_token(
    client: QuickBooks,
    config: Dict[str, str],
    token_path: Path = TOKEN_FILE,
) -> bool:
    """Refresh OAuth2 access token using the provided client.

    Updates ``config`` with the new tokens and saves them to ``token_path``.
    Returns ``True`` on success, ``False`` otherwise.
    """
    try:
        print("⏳ Refreshing access token...")
        client.auth_client.refresh()

        # Sync tokens on the QuickBooks client itself
        client.access_token = client.auth_client.access_token
        client.refresh_token = client.auth_client.refresh_token

        config["ACCESS_TOKEN"] = client.auth_client.access_token
        config["REFRESH_TOKEN"] = client.auth_client.refresh_token

        token_data = {
            "access_token": client.auth_client.access_token,
            "refresh_token": client.auth_client.refresh_token,
            "realm_id": config.get("REALM_ID"),
        }
        with open(token_path, "w", encoding="utf-8") as f:
            json.dump(token_data, f, indent=4)
        print("💾 Tokens have been refreshed and saved to qb_tokens.json.")
        prompt = "upload the 'qb_tokens.json' file to SharePoint, once done, enter 'ok'"
        while True:
            response = input(f"{prompt}\n").strip().lower()
            if response == "ok":
                break
        return True
    except Exception as e:
        print(f"❌ CRITICAL: Failed to refresh access token: {e}")
        print("   Please re-authenticate via the OAuth Playground and update qb_tokens.json.")
        return False
