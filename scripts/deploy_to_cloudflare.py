#!/usr/bin/env python3
"""
Deploy DSR Dashboard HTML to Cloudflare Pages via wrangler.

Usage:
    python3 deploy_to_cloudflare.py <path-to-dashboard.html>

Pipeline: HTML → wrangler pages deploy → live at dsr-26s.pages.dev
(GitHub integration is intentionally disconnected — deploys go direct.)
"""

import sys
import os
import json
import subprocess
import shutil
import tempfile
from pathlib import Path
from datetime import datetime

# ── Configuration ──────────────────────────────────────────────
CLOUDFLARE_URL = "https://dsr-26s.pages.dev"
PROJECT_NAME = "dsr"
CONFIG_FILENAME = ".deploy-config.json"
WRANGLER_INSTALL_DIR = "/tmp/npm-wrangler"


def run(cmd, cwd=None, check=True, env=None):
    result = subprocess.run(
        cmd, shell=True, cwd=cwd,
        capture_output=True, text=True, env=env
    )
    if check and result.returncode != 0:
        print(f"ERROR running: {cmd}")
        print(result.stderr[-1000:] if result.stderr else "(no stderr)")
        sys.exit(1)
    return result


def load_config():
    search_paths = [
        Path(__file__).parent,
        Path(__file__).parent.parent,
        Path.cwd(),
    ]
    for base in search_paths:
        cfg_path = base / CONFIG_FILENAME
        if cfg_path.exists():
            try:
                return json.loads(cfg_path.read_text(encoding='utf-8'))
            except (json.JSONDecodeError, OSError):
                continue
    return {}


def get_api_token(cfg):
    token = os.environ.get("CLOUDFLARE_API_TOKEN", "")
    if token:
        return token
    token = cfg.get("cloudflare_api_token", "")
    if token:
        return token
    print("ERROR: CLOUDFLARE_API_TOKEN not found.")
    print(f"Set it with: export CLOUDFLARE_API_TOKEN=cfut_...")
    print(f"Or add cloudflare_api_token to {CONFIG_FILENAME}")
    sys.exit(1)


def get_account_id(cfg):
    account_id = os.environ.get("CLOUDFLARE_ACCOUNT_ID", "")
    if account_id:
        return account_id
    return cfg.get("cloudflare_account_id", "")


def ensure_wrangler():
    """Return path to wrangler binary, installing if needed."""
    # Check if already installed in our temp dir
    wrangler_bin = Path(WRANGLER_INSTALL_DIR) / "bin" / "wrangler"
    if wrangler_bin.exists():
        return str(wrangler_bin)

    # Check system wrangler
    r = run("which wrangler", check=False)
    if r.returncode == 0:
        return r.stdout.strip()

    # Install to temp location
    print("Installing wrangler (one-time)...")
    run(f"npm install -g wrangler --prefix {WRANGLER_INSTALL_DIR}")
    if not wrangler_bin.exists():
        print("ERROR: wrangler install failed")
        sys.exit(1)
    return str(wrangler_bin)


def deploy(html_path: str):
    """Deploy an HTML file to Cloudflare Pages via wrangler."""
    html_file = Path(html_path)
    if not html_file.exists():
        print(f"ERROR: File not found: {html_path}")
        sys.exit(1)

    cfg = load_config()
    api_token = get_api_token(cfg)
    account_id = get_account_id(cfg)
    wrangler = ensure_wrangler()

    # Stage in a temp dir (wrangler deploys a directory, not a single file)
    deploy_dir = Path(tempfile.mkdtemp(prefix="dsr-deploy-"))
    try:
        dest = deploy_dir / "index.html"
        shutil.copy2(str(html_file), str(dest))
        print(f"Staging: {html_file.name} → {dest}")

        # Build env for wrangler
        env = os.environ.copy()
        env["CLOUDFLARE_API_TOKEN"] = api_token
        if account_id:
            env["CLOUDFLARE_ACCOUNT_ID"] = account_id

        today = datetime.now().strftime("%Y-%m-%d")
        print(f"Deploying to Cloudflare Pages ({PROJECT_NAME})...")
        r = run(
            f'{wrangler} pages deploy {deploy_dir} '
            f'--project-name={PROJECT_NAME} '
            f'--branch=main '
            f'--commit-message="DSR dashboard {today}"',
            env=env
        )

        output = r.stdout + r.stderr
        print(output.strip())
        print(f"\n✓ Deployed successfully!")
        print(f"  Live at: {CLOUDFLARE_URL}")

    finally:
        shutil.rmtree(str(deploy_dir), ignore_errors=True)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 deploy_to_cloudflare.py <path-to-dashboard.html>")
        sys.exit(1)

    deploy(sys.argv[1])
