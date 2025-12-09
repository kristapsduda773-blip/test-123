#!/usr/bin/env python3
"""
Fetch the specified Entra ID (Azure AD) groups from Microsoft Graph and
classify each one as either cloud-only or on-premises (hybrid) based on
their on-premises attributes.

Usage:
    export AZURE_TENANT_ID="..."
    export AZURE_CLIENT_ID="..."
    export AZURE_CLIENT_SECRET="..."
    python group_source_report.py --output groups.csv

By default the script looks up the group names listed in DEFAULT_GROUP_NAMES.
Pass --names-file to provide a custom newline-separated list instead.
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

try:
    import msal  # type: ignore
except ImportError as exc:  # pragma: no cover - dependency hint
    raise SystemExit(
        "Missing dependency 'msal'. Install it with 'pip install msal requests'."
    ) from exc

try:
    import requests
except ImportError as exc:  # pragma: no cover - dependency hint
    raise SystemExit(
        "Missing dependency 'requests'. Install it with 'pip install requests'."
    ) from exc


GRAPH_SCOPE = "https://graph.microsoft.com/.default"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_TIMEOUT = 15

DEFAULT_GROUP_NAMES: Tuple[str, ...] = (
    "DEV-ATD",
    "DEV-AVIS-cloud",
    "DEV-BDAS-cloud",
    "DEV-CAKEHR-cloud",
    "DEV-DELTAPV-cloud",
    "DEV-ESTAPIKS2-cloud",
    "DEV-Evo-Roads",
    "DEV-FITS-cloud",
    "DEV-INTRANET-cloud",
    "DEV-LAMBDAPV-cloud",
    "DEV-MANAGEMENT-cloud",
    "DEV-NILDA2-cloud",
    "DEV-OPVS–CargoRail",
    "DEV-PRESERVICA-cloud",
    "DEV-VADDVS",
    "Dots-Sales",
    "Product Group",
    "BW-DEV-ATD",
    "BW-DEV-Common",
    "BW-DEV-EvoRoads",
    "BW-DEV-Kappa",
    "BW-DEV-LDz-OPVS",
    "BW-DEV-SAGE-MAGS",
    "DEV-AIHEN-cloud",
    "DEV-AIROS",
    "DEV-Digitalizacija",
    "DEV-EXT-AKKA-LAA",
    "DEV-External-ATD-SMARTIN",
    "DEV-External-SMARTIN-DESIGN",
    "DEV-EXT-Estapiks2",
    "DEV-EXT-Fits",
    "DEV-EXT-KAMIS",
    "DEV-EXT-Peruza",
    "DEV-EXT-Preservica",
    "DEV-FITS",
    "DEV-IC-FITS",
    "SQL-00-PBI-Sync",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Classify configured Entra ID groups as cloud-only or on-prem."
    )
    parser.add_argument(
        "--names-file",
        type=Path,
        help="Path to a newline-separated file with custom group names.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Optional CSV file to store the detailed results.",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT,
        help=f"HTTP timeout for Graph calls (seconds, default {DEFAULT_TIMEOUT}).",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print Graph query diagnostics to stderr.",
    )
    return parser.parse_args()


def load_group_names(names_file: Optional[Path]) -> Sequence[str]:
    if not names_file:
        return DEFAULT_GROUP_NAMES

    if not names_file.exists():
        raise SystemExit(f"Names file '{names_file}' was not found.")

    contents = names_file.read_text(encoding="utf-8").splitlines()
    cleaned = [line.strip() for line in contents if line.strip() and not line.startswith("#")]

    if not cleaned:
        raise SystemExit(f"No valid group names were found in '{names_file}'.")

    return cleaned


def require_env(var_name: str) -> str:
    value = os.getenv(var_name)
    if not value:
        raise SystemExit(
            f"Environment variable '{var_name}' is required but missing. "
            "Set it and retry."
        )
    return value


def build_confidential_client() -> msal.ConfidentialClientApplication:
    tenant_id = require_env("AZURE_TENANT_ID")
    client_id = require_env("AZURE_CLIENT_ID")
    client_secret = require_env("AZURE_CLIENT_SECRET")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    return msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=authority,
        client_credential=client_secret,
    )


def acquire_graph_token(app: msal.ConfidentialClientApplication) -> str:
    result = app.acquire_token_silent([GRAPH_SCOPE], account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=[GRAPH_SCOPE])

    if "access_token" not in result:
        raise SystemExit(
            f"Failed to obtain Graph token: {result.get('error_description', 'unknown error')}"
        )

    return result["access_token"]


def debug(message: str, enabled: bool) -> None:
    if enabled:
        print(message, file=sys.stderr)


def fetch_group(
    access_token: str, display_name: str, timeout: int, verbose: bool
) -> Optional[Dict[str, Any]]:
    filter_value = display_name.replace("'", "''")
    params = {
        "$filter": f"displayName eq '{filter_value}'",
        "$select": (
            "id,displayName,onPremisesSyncEnabled,onPremisesDomainName,"
            "onPremisesSecurityIdentifier,securityEnabled,mailEnabled,groupTypes,createdDateTime"
        ),
        "$count": "true",
    }
    headers = {
        "Authorization": f"Bearer {access_token}",
        "ConsistencyLevel": "eventual",
    }

    debug(f"Querying Graph for '{display_name}'", verbose)
    response = requests.get(
        f"{GRAPH_BASE_URL}/groups",
        headers=headers,
        params=params,
        timeout=timeout,
    )
    response.raise_for_status()
    payload = response.json()
    matches = payload.get("value", [])

    exact = next(
        (item for item in matches if item.get("displayName", "").lower() == display_name.lower()),
        None,
    )
    if exact:
        return exact

    return matches[0] if matches else None


def classify_group(group: Optional[Dict[str, Any]]) -> Tuple[str, str]:
    if not group:
        return "not-found", "Group was not returned by Microsoft Graph"

    sync_flag = group.get("onPremisesSyncEnabled")
    domain = group.get("onPremisesDomainName")
    onprem_sid = group.get("onPremisesSecurityIdentifier")

    if sync_flag is True:
        return "on-prem", "onPremisesSyncEnabled=True"
    if domain:
        return "on-prem", "onPremisesDomainName populated"
    if onprem_sid:
        return "on-prem", "onPremisesSecurityIdentifier populated"
    if sync_flag is False:
        return "cloud", "onPremisesSyncEnabled=False"

    return "cloud", "No on-premises attributes detected"


def evaluate_groups(
    group_names: Sequence[str], access_token: str, timeout: int, verbose: bool
) -> List[Dict[str, str]]:
    records: List[Dict[str, str]] = []
    for name in group_names:
        try:
            group = fetch_group(access_token, name, timeout, verbose)
            source, reason = classify_group(group)
            records.append(
                {
                    "requested_name": name,
                    "matched_name": group.get("displayName", "") if group else "",
                    "source": source,
                    "reason": reason,
                    "group_id": group.get("id", "") if group else "",
                    "domain": group.get("onPremisesDomainName", "") if group else "",
                }
            )
        except requests.RequestException as exc:
            records.append(
                {
                    "requested_name": name,
                    "matched_name": "",
                    "source": "error",
                    "reason": f"Graph request failed: {exc}",
                    "group_id": "",
                    "domain": "",
                }
            )
    return records


def print_table(records: Sequence[Dict[str, str]]) -> None:
    if not records:
        print("No records to display.")
        return

    columns = (
        ("requested_name", "Requested"),
        ("matched_name", "Graph Match"),
        ("source", "Source"),
        ("domain", "On-Prem Domain"),
        ("reason", "Reason"),
        ("group_id", "Object Id"),
    )

    widths = {
        key: max(len(header), *(len(record.get(key, "")) for record in records))
        for key, header in columns
    }

    header_row = "  ".join(header.ljust(widths[key]) for key, header in columns)
    separator = "  ".join("-" * widths[key] for key, _ in columns)
    print(header_row)
    print(separator)

    for record in records:
        print(
            "  ".join(record.get(key, "").ljust(widths[key]) for key, _ in columns)
        )


def write_csv(path: Path, records: Sequence[Dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = ["requested_name", "matched_name", "source", "reason", "group_id", "domain"]
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(records)
    print(f"Saved CSV results to '{path}'.")


def main() -> None:
    args = parse_args()
    names = load_group_names(args.names_file)
    app = build_confidential_client()
    token = acquire_graph_token(app)
    records = evaluate_groups(names, token, args.timeout, args.verbose)
    print_table(records)
    if args.output:
        write_csv(args.output, records)


if __name__ == "__main__":
    main()
