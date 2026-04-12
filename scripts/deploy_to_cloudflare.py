#!/usr/bin/env python3
"""
Deploy DSR Dashboard HTML to Cloudflare Pages via GitHub.

Usage:
    python3 deploy_to_cloudflare.py <path-to-dashboard.html>

Pipeline: HTML → GitHub push → Cloudflare Pages auto-deploy → live at dsr-26s.pages.dev
"""

import sys
import os
import json
import subprocess
import shutil
from pathlib import Path
from datetime import datetime

# ── Configuration ──────────────────────────────────────────────
GITHUB_REPO = "pawitwayachut/dsr-dashboard"
GITHUB_BRANCH = "main"
CLOUDFLARE_URL = "https://dsr-26s.pages.dev"
DEPLOY_DIR = "/tmp/dsr-deploy"
CONFIG_FILENAME = ".deploy-config.json"


def run(cmd, cwd=None, check=True):
    """Run a shell command and return output."""
    result = subprocess.run(
        cmd, shell=True, cwd=cwd,
        capture_output=True, text=True
    )
    if check and result.returncode != 0:
        print(f"ERROR: {cmd}\n{result.stderr}")
        sys.exit(1)
    return result


def get_github_token():
    """Get GitHub token from env var, or fall back to .deploy-config.json."""
    token = os.environ.get("GITHUB_TOKEN", "")
    if token:
        return token

    # Search for config file: script dir, then parent, then cwd
    search_paths = [
        Path(__file__).parent,
        Path(__file__).parent.parent,
        Path.cwd(),
    ]
    for base in search_paths:
        cfg_path = base / CONFIG_FILENAME
        if cfg_path.exists():
            try:
                cfg = json.loads(cfg_path.read_text(encoding='utf-8'))
                token = cfg.get("github_token", "")
                if token:
                    print(f"  Using token from {cfg_path}")
                    return token
            except (json.JSONDecodeError, OSError):
                continue

    print("ERROR: GITHUB_TOKEN not found.")
    print("Set it with: export GITHUB_TOKEN=ghp_yourtoken")
    print(f"Or create {CONFIG_FILENAME} with {{\"github_token\": \"...\"}}")
    sys.exit(1)


def deploy(html_path: str):
    """Deploy an HTML file to Cloudflare Pages via GitHub."""
    html_file = Path(html_path)
    if not html_file.exists():
        print(f"ERROR: File not found: {html_path}")
        sys.exit(1)

    token = get_github_token()
    repo_url = f"https://pawitwayachut:{token}@github.com/{GITHUB_REPO}.git"
    deploy_dir = Path(DEPLOY_DIR)

    # Clone or update repo
    if deploy_dir.exists():
        print("Updating existing repo...")
        run("git fetch origin && git reset --hard origin/main", cwd=str(deploy_dir))
    else:
        print("Cloning repo...")
        run(f"git clone {repo_url} {deploy_dir}")

    # Configure git
    run('git config user.email "pawit.wayachut@gmail.com"', cwd=str(deploy_dir))
    run('git config user.name "Pawit Wayachut"', cwd=str(deploy_dir))

    # Copy HTML as index.html
    dest = deploy_dir / "index.html"
    shutil.copy2(str(html_file), str(dest))
    print(f"Copied {html_file.name} → index.html")

    # Check if there are changes
    status = run("git status --porcelain", cwd=str(deploy_dir))
    if not status.stdout.strip():
        print("No changes to deploy.")
        print(f"Live at: {CLOUDFLARE_URL}")
        return

    # Commit and push
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    commit_msg = f"Update DSR dashboard — {timestamp}"

    run("git add index.html", cwd=str(deploy_dir))
    run(f'git commit -m "{commit_msg}"', cwd=str(deploy_dir))
    run("git push origin main", cwd=str(deploy_dir))

    print(f"\n✓ Deployed successfully!")
    print(f"  Live at: {CLOUDFLARE_URL}")
    print(f"  Commit: {commit_msg}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 deploy_to_cloudflare.py <path-to-dashboard.html>")
        sys.exit(1)

    deploy(sys.argv[1])
