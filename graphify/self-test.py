"""Negative fixtures for the graphify publishing tooling (R2-G01 onwards, §13.10).

Every guard in this directory exists to refuse something. A guard is only worth having if
the refusal is exercised, so each case here drives the real code path and asserts that it
fails; the passing case is included too, because a guard that refuses everything is just as
broken as one that refuses nothing.

Dangerous targets are checked through the pure guard function rather than by running the
staging script, so a regression in the guard cannot delete the directory the fixture names.

Usage:
    python graphify/self-test.py
"""
from __future__ import annotations

import fnmatch
import hashlib
import importlib.util
import json
import os

NL = chr(10)
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parent

failures: list[str] = []
checks = 0


def load(module_name: str, filename: str):
    """Imports a sibling script whose filename is not a valid module name."""
    spec = importlib.util.spec_from_file_location(module_name, HERE / filename)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def check(description: str, condition: bool) -> None:
    """Records one fixture outcome."""
    global checks
    checks += 1
    if condition:
        print(f"  pass  {description}")
    else:
        print(f"  FAIL  {description}")
        failures.append(description)


def refuses(fn, description: str) -> None:
    """Asserts that a guard raises SystemExit for the given target."""
    try:
        fn()
    except SystemExit:
        check(description, True)
        return
    check(description, False)


def allows(fn, description: str) -> None:
    """Asserts that a guard accepts a legitimate target."""
    try:
        fn()
    except SystemExit as exc:
        print(f"  FAIL  {description} (refused: {exc})")
        failures.append(description)
        global checks
        checks += 1
        return
    check(description, True)


def write_marker(stage, directory, **overrides):
    """Writes a marker of the shape the staging script itself produces."""
    marker = {
        "schema": stage.MARKER_SCHEMA,
        "tool": "graphify/stage-corpus.py",
        "repository": stage.repository_identity(),
        "nonce": "0123456789abcdef0123456789abcdef",
    }
    marker.update(overrides)
    (directory / stage.STAGING_MARKER).write_text(json.dumps(marker), encoding="utf-8")


def write_manifest(directory, files, run_id="0123456789abcdef0123456789abcdef"):
    """Writes a corpus manifest describing exactly the named files.

    <param name="directory">The staging directory.</param>
    <param name="files">Directory-relative POSIX paths the manifest claims.</param>
    <param name="run_id">Run id, which has to match the marker's nonce.</param>
    """
    (directory / "corpus-manifest.json").write_text(json.dumps({
        "run_id": run_id,
        "files": {
            name: hashlib.sha256((directory / name).read_bytes()).hexdigest()
            if (directory / name).is_file() else "0" * 64
            for name in files
        },
        "file_count": len(files),
        "corpus_sha256": "0" * 64,
        "allowlist_sha256": "0" * 64,
    }), encoding="utf-8")


def stage_a_directory(stage, root, name, files=("Program.cs",)):
    """Builds a directory that looks exactly like one this tool staged.

    <param name="stage">The loaded staging module.</param>
    <param name="root">Where to create it.</param>
    <param name="name">Directory name.</param>
    <param name="files">Files to create and record.</param>
    <returns>The directory.</returns>
    """
    directory = root / name
    directory.mkdir(parents=True, exist_ok=True)
    for relative in files:
        target = directory / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("//", encoding="utf-8")
    write_marker(stage, directory)
    write_manifest(directory, files)
    return directory


def test_staging_guard() -> None:
    """R2-G01/R8-G03: staging only ever creates a new directory, and deletes nothing."""
    print("R2-G01 staging target guard")
    stage = load("stage_corpus", "stage-corpus.py")
    guard = stage.ensure_safe_staging_target

    refuses(lambda: guard(REPO_ROOT), "repository root is refused")
    refuses(lambda: guard(REPO_ROOT / "Tools"), "Tools/ inside the repo is refused")
    refuses(lambda: guard(REPO_ROOT / ".git"), ".git/ inside the repo is refused")
    refuses(lambda: guard(REPO_ROOT / "docs"), "docs/ inside the repo is refused")
    refuses(lambda: guard(REPO_ROOT.parent), "a parent of the repo is refused")

    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp).resolve()

        unrelated = root / "someone-elses-data"
        unrelated.mkdir()
        (unrelated / "notes.txt").write_text("keep me", encoding="utf-8")
        refuses(lambda: guard(unrelated),
                "an unrelated non-empty directory is refused")

        staged = stage_a_directory(stage, root, "staging")
        refuses(lambda: guard(staged),
                "a directory this tool staged before is refused too")

        allows(lambda: guard(root / "brand-new"), "a directory that does not exist yet is allowed")

        # R3-G01: a marker is only honoured when its contents say it belongs here.
        forged = stage_a_directory(stage, root, "forged")

        (forged / stage.STAGING_MARKER).write_text("", encoding="utf-8")
        refuses(lambda: guard(forged), "an empty marker file is refused")

        (forged / stage.STAGING_MARKER).write_text("not json", encoding="utf-8")
        refuses(lambda: guard(forged), "a non-JSON marker is refused")

        write_marker(stage, forged, repository="some-other-repository")
        refuses(lambda: guard(forged), "a marker from another repository is refused")

        write_marker(stage, forged, schema=stage.MARKER_SCHEMA + 1)
        refuses(lambda: guard(forged), "a marker with an unknown schema is refused")

        write_marker(stage, forged, nonce="short")
        refuses(lambda: guard(forged), "a marker without a real nonce is refused")

        write_marker(stage, forged)
        refuses(lambda: guard(forged),
                "even a well-formed marker cannot authorise deleting the directory it sits in")

        # R4-G01: a marker proves this tool created the directory. It does not prove the directory
        # is still only that.
        intruded = stage_a_directory(stage, root, "intruded")
        (intruded / "personal-notes.txt").write_text("someone else's work", encoding="utf-8")
        refuses(lambda: guard(intruded),
                "a staged directory holding a file the manifest does not describe is refused")

        nested_intruder = stage_a_directory(stage, root, "nested-intruder")
        (nested_intruder / "sub").mkdir()
        (nested_intruder / "sub" / "notes.md").write_text("mine", encoding="utf-8")
        refuses(lambda: guard(nested_intruder),
                "an unlisted file in a subdirectory is refused too")

        work_dir = stage_a_directory(stage, root, "with-work-dir")
        (work_dir / "graphify-out").mkdir()
        (work_dir / "graphify-out" / "graph.json").write_text("{}", encoding="utf-8")
        refuses(lambda: guard(work_dir),
                "an extraction output the work directory has not recorded is refused")
        load("corpus_allowlist_guard", "corpus_allowlist.py").write_work_inventory(
            work_dir, {"graphify-out/graph.json": "graph"}, "0123456789abcdef0123456789abcdef")
        refuses(lambda: guard(work_dir),
                "a recorded work directory still cannot authorise its own deletion")

        no_manifest = stage_a_directory(stage, root, "no-manifest")
        (no_manifest / "corpus-manifest.json").unlink()
        refuses(lambda: guard(no_manifest),
                "a marker with no manifest beside it is refused")

        other_run = stage_a_directory(stage, root, "other-run")
        write_manifest(other_run, ["Program.cs"], run_id="f" * 32)
        refuses(lambda: guard(other_run),
                "a marker and manifest from different runs are refused")

        empty_manifest = stage_a_directory(stage, root, "empty-manifest")
        write_manifest(empty_manifest, [])
        refuses(lambda: guard(empty_manifest),
                "a manifest describing no files is refused")

        empty = root / "empty"
        empty.mkdir()
        allows(lambda: guard(empty), "an existing empty directory is allowed")


def test_staged_manifest_binds_bytes() -> None:
    """R2-G03: the manifest must describe the staged copy, not the working tree."""
    print("R2-G03 manifest binds the staged bytes")
    allowlist = load("corpus_allowlist_selftest", "corpus_allowlist.py")

    with tempfile.TemporaryDirectory() as tmp:
        staging = Path(tmp).resolve() / "corpus"
        result = subprocess.run(
            [sys.executable, str(HERE / "stage-corpus.py"), "--out", str(staging)],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        check("staging a fresh directory succeeds", result.returncode == 0)
        if result.returncode != 0:
            print(result.stdout + result.stderr)
            return

        stage = load("stage_corpus_marker", "stage-corpus.py")
        check("the marker is written", (staging / stage.STAGING_MARKER).is_file())

        manifest = json.loads((staging / "corpus-manifest.json").read_text(encoding="utf-8"))
        recomputed = allowlist.build_manifest(staging)
        check("the recorded digest matches the staged copy",
              recomputed["corpus_sha256"] == manifest["corpus_sha256"])

        victim = staging / "Program.cs"
        original = victim.read_bytes()
        victim.write_bytes(original + b"\n// tampered\n")
        tampered = allowlist.build_manifest(staging)
        check("editing a staged file changes the recomputed digest",
              tampered["corpus_sha256"] != manifest["corpus_sha256"])
        victim.write_bytes(original)

        # R3-G02: build_manifest only walks what the allowlist selects, so a file added to the
        # staging tree afterwards was invisible to both the digest and the added/missing diff.
        builder = load("build_public_map_selftest", "build-public-map.py")
        intruder = staging / "secret.md"
        intruder.write_text("not part of the corpus", encoding="utf-8")
        refuses(lambda: builder.verify_corpus_manifest(staging),
                "a file the manifest does not describe is refused")
        intruder.unlink()

        nested = staging / "Core" / "planted.cs"
        nested.write_text("// planted", encoding="utf-8")
        refuses(lambda: builder.verify_corpus_manifest(staging),
                "a planted file inside an allowlisted directory is refused")
        nested.unlink()

        allows(lambda: builder.verify_corpus_manifest(staging),
               "the untouched staging directory still verifies")

        # R3-G03: the run id is what ties the extraction and the page to this staging run.
        manifest_path = staging / "corpus-manifest.json"
        recorded = json.loads(manifest_path.read_text(encoding="utf-8"))
        check("staging records a run id", len(str(recorded.get("run_id", ""))) == 32)

        without_run_id = dict(recorded)
        without_run_id.pop("run_id", None)
        manifest_path.write_text(json.dumps(without_run_id), encoding="utf-8")
        refuses(lambda: builder.verify_corpus_manifest(staging),
                "a manifest with no staging run id is refused")
        manifest_path.write_text(json.dumps(recorded), encoding="utf-8")

        shutil.rmtree(staging)


def test_artifacts_share_one_run() -> None:
    """R3-G03: the extraction, the graph and the page must come from one corpus run."""
    print("R3-G03 artifacts are bound to one run")
    builder = load("build_public_map_runs", "build-public-map.py")
    bind = builder.ensure_artifacts_share_one_run

    def page(ids: list[str]) -> str:
        drawn = [{"id": i, "label": i} for i in ids]
        return "const RAW_NODES = " + json.dumps(drawn) + ";" + chr(10) + "const RAW_EDGES = [];"

    # The node sets have to match exactly: a graph that drops one of the extraction's nodes is
    # itself from another run, which the R4-G02 fixtures below cover.
    extraction = {"nodes": [{"id": "a"}, {"id": "b"}]}
    graph = {"nodes": [{"id": "a", "community": "0"}, {"id": "b", "community": "1"}]}

    allows(lambda: bind(extraction, graph, page(["0", "1"])),
           "a page drawing this graph's communities is accepted")
    allows(lambda: bind(extraction, graph, page(["a", "b"])),
           "a page drawing this graph's own nodes is accepted")

    other_run = {"nodes": [{"id": "x", "community": "0"}, {"id": "y", "community": "1"}]}
    refuses(lambda: bind(extraction, other_run, page(["0", "1"])),
            "a graph from another run's extraction is refused")
    refuses(lambda: bind(extraction, graph, page(["7", "8"])),
            "a page exported from another run's graph is refused")
    refuses(lambda: bind(extraction, graph, "const RAW_EDGES = [];"),
            "a page carrying no node array is refused")



MUST_REFUSE = (
    ("Windows drive path", "D:" + chr(92) + "GIT" + chr(92) + "JaJa" + chr(92) + "server"),
    # R23-G03: GitHub fine-grained tokens, `github_pat_` + 22 + `_` + 59 characters.
    ("GitHub fine-grained token", "token = github_pat_" + "A1b2C3d4E5" * 2 + "Ab" + "_" + "Zz9y8X7w6V" * 5 + "Q1w2E3r4t"),
    ("UNC share", chr(92) * 2 + "fileserver" + chr(92) + "share" + chr(92) + "build"),
    ("Windows device path", chr(92) * 2 + "?" + chr(92) + "C:" + chr(92) + "long"),
    ("POSIX home", "/home/jaja/src/main.cs"),
    ("macOS home", "/Users/jaja/src/main.cs"),
    ("system configuration path", "/etc/passwd"),
    ("system data path", "/var/log/build.log"),
    ("optional software path", "/opt/aspose/lib.so"),
    ("loopback name", "listening on localhost:5000"),
    ("loopback address", "http://127.0.0.1:8080/"),
    ("wildcard address", "bind 0.0.0.0"),
    ("RFC 1918 /8", "10.1.2.3"),
    ("RFC 1918 /16", "192.168.1.50"),
    ("RFC 1918 /12 lower bound", "172.16.4.9"),
    ("RFC 1918 /12 upper bound", "172.31.255.1"),
    ("cloud metadata address", "169.254.169.254"),
    ("carrier-grade NAT address", "100.100.5.7"),
    ("IPv6 loopback", "connect to ::1 now"),
    ("IPv6 unique local address", "fd12:3456:789a::1"),
    ("IPv6 link-local address", "fe80::1ff:fe23:4567"),
    ("internal hostname", "build01.corp.internal"),
    ("secret-shaped string", "api_key = abcdefg"),
    ("e-mail address", "jaja@example.com"),
    # R19-G01. Every one of these returned zero findings: the keyword pattern wants the
    # delimiter adjacent, and a closing quote, a header name or a directory sits in between.
    ("quoted JSON secret", '{"api_key": "abcdefghijkl"}'),
    ("quoted JSON password", '{"password": "hunter2hunter2"}'),
    ("quoted YAML secret", "  secret_key: 'abcdefghijkl'"),
    ("authorization header", "Authorization: Bearer abcdefghijklmnopqrst"),
    ("bare bearer credential", "send bearer abcdefghijklmnopqrstuvwx"),
    ("docker secret mount", "/run/secrets/github_token"),
    ("kubernetes serviceaccount token", "/var/run/secrets/kubernetes.io/serviceaccount/token"),
    ("workflow secret reference", "${{ secrets.GITHUB_TOKEN }}"),
    ("AWS access key id", "AKIAIOSFODNN7EXAMPLE"),
    ("Google API key", "AIzaSyA1234567890abcdefghijklmnopqrstuv"),
    ("Slack token", "xoxb-1234567890-abcdefghij"),
    ("GitLab personal access token", "glpat-abcdefghijklmnopqrst"),
    # R20-G01: three the direct probe walked straight through.
    ("quoted client secret", '{"client_secret": "abcdefghijkl"}'),
    ("AWS secret access key assignment", "aws_secret_access_key = wJalrXUtnFEMIK7MDENG"),
    ("short quoted password", '{"password": "hunt3r"}'),
    # R21-G01: every one of these passed; the header never has `:` or `=` after the keyword.
    ("PKCS#8 private key header", "-----BEGIN PRIVATE KEY-----"),
    ("RSA private key header", "-----BEGIN RSA PRIVATE KEY-----"),
    ("EC private key header", "-----BEGIN EC PRIVATE KEY-----"),
    ("DSA private key header", "-----BEGIN DSA PRIVATE KEY-----"),
    ("OpenSSH private key header", "-----BEGIN OPENSSH PRIVATE KEY-----"),
    ("encrypted private key header", "-----BEGIN ENCRYPTED PRIVATE KEY-----"),
    ("private key footer with CRLF and indent", "  -----END RSA PRIVATE KEY-----\r\n"),
    # R21-G02: placeholders are refused by policy, not excused by shape.
    ("credential-named key with an angle placeholder", '{"client_secret": "<your value>"}'),
    ("credential-named key with an env placeholder", '{"client_secret": "${CLIENT_SECRET}"}'),
    ("unquoted credential with a brace placeholder", "client_secret = {placeholder}"),
)

MUST_ALLOW = (
    ("a repository-relative source path", "Core/Session/DocumentSession.cs"),
    ("a site-relative link", "/architecture-map/index.html"),
    ("the published site", "https://xjustloveux.github.io/aspose-mcp-server/"),
    ("the container registry", "https://ghcr.io/xjustloveux/aspose-mcp-server"),
    ("vendor documentation", "https://docs.aspose.com/words/net/"),
    ("an XML namespace", "https://www.w3.org/2000/svg"),
    ("hex colours that look like an IPv6 prefix", "colour #fd7e14 and #10a37f"),
    ("public addresses either side of RFC 1918", "172.15.1.1 and 172.32.1.1"),
    ("public addresses either side of the NAT range", "100.63.0.1 and 100.128.0.1"),
    ("a project file name", "AsposeMcpServer.csproj"),
    ("version numbers", "version 1.2.0 build 10.0.10"),
    # The lookalikes the widened patterns must still let through. This project's published
    # documentation is largely a list of option names, and every one of them contains a word
    # the patterns above look for.
    ("an API key environment variable name", "<code>ASPOSE_AUTH_APIKEY_KEYS</code>"),
    ("an API key command-line flag", "--auth-apikey-header name"),
    ("an API key header name", "<td>X-API-Key</td>"),
    ("a documented but unset config value", '{"apiKey": ""}'),
    ("a config key with no value", '"password":'),
    ("the word bearer in prose", "the bearer of this token is checked"),
    ("the word secret in prose", "secret management is a deployment concern"),
    ("a docs path that mentions secrets", "/architecture-map/secrets-overview.html"),
    ("an authorization mode name", "authorization mode: gateway"),
    ("a base64 word that is not a vendor key", "AKIAISNOTAKEYATALL is prose"),
    ("an empty quoted password", '{"password": ""}'),
    ("a key that merely contains the word", '{"password_policy": "rotate quarterly"}'),
    # R21-G01 controls: public material is not a private key.
    ("a public key header", "-----BEGIN PUBLIC KEY-----"),
    ("a certificate header", "-----BEGIN CERTIFICATE-----"),
    ("prose that mentions a private key", "keep your private key somewhere safe"),
)

SCANNER_DRIVER = """
Import-Module (Join-Path '{module_dir}' 'LeakScanner.psm1') -Force
$samples = Get-Content -Raw -LiteralPath '{samples}' | ConvertFrom-Json
$results = foreach ($sample in $samples) {{
    @(@(Get-LeakFinding -Text $sample).Count)
}}
Set-Content -LiteralPath '{results}' -Value (@($results) | ConvertTo-Json -Compress) -Encoding utf8
"""


SCOPED_DRIVER = """
Import-Module (Join-Path '{module_dir}' 'LeakScanner.psm1') -Force
$samples = Get-Content -Raw -LiteralPath '{samples}' | ConvertFrom-Json
$results = foreach ($sample in $samples) {{
    @(@(Get-LeakFinding -Text $sample.text -Scope $sample.scope).Count)
}}
Set-Content -LiteralPath '{results}' -Value (@($results) | ConvertTo-Json -Compress) -Encoding utf8
"""

# R20-G01, second pass: what each scope must refuse and must let through.
SCOPED = (
    ("Generated", "a loopback host in the generated map is refused", "listening on localhost:5000", True),
    ("Authored", "a page telling the reader to connect to localhost is allowed", "connect to http://localhost:3000/mcp", False),
    ("Authored", "a page telling the reader to bind 0.0.0.0 is allowed", "set the host to 0.0.0.0 to listen on all interfaces", False),
    ("Authored", "a private address in prose is still refused", "the build runs on 192.168.1.50", True),
    ("Authored", "a workspace root in prose is refused", "open D:" + chr(92) + "GIT" + chr(92) + "JaJa", True),
    ("Authored", "a home directory in prose is refused", "see C:" + chr(92) + "Users" + chr(92) + "jaja" + chr(92) + "Documents", True),
    ("Authored", "a POSIX home directory in prose is refused", "edit /home/jaja/.config/app.json", True),
    ("Authored", "an install location in a how-to is allowed", "extract to C:" + chr(92) + "Tools" + chr(92) + "aspose-mcp-server" + chr(92), False),
    ("Authored", "a home placeholder in a how-to is allowed", "config lives under C:" + chr(92) + "Users" + chr(92) + "...", False),
    ("Authored", "a system directory in a how-to is allowed", "copy fonts to /usr/share/fonts/custom", False),
    ("Authored", "a temp directory in a how-to is allowed", "snapshots go to /tmp/extensions/snapshot_xxx.pdf", False),
    ("Authored", "a credential in prose is still refused", '{"api_key": "abcdefghijkl"}', True),
    ("Binary", "three bytes that happen to spell a drive letter are allowed", "GIF89a\u00e5E:" + chr(92) + "\u0001\u00ff", False),
    ("Binary", "a real path embedded as ASCII is refused", "Comment: E:" + chr(92) + "Users" + chr(92) + "jaja" + chr(92) + "demo.gif", True),
    ("Binary", "a vendor key embedded as ASCII is refused", "\u0001AKIAIOSFODNN7EXAMPLE\u0002", True),
    ("Binary", "a loopback host in binary is not looked for", "\u00ff localhost \u00ff", False),
)


def test_leak_scanner_scopes() -> None:
    """R20-G01: the rules a file is held to depend on what kind of file it is."""
    print("R20-G01 leak scanner scopes")
    samples = [{"scope": scope, "text": text} for scope, _, text, _ in SCOPED]
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp).resolve()
        samples_path = root / "samples.json"
        results_path = root / "results.json"
        driver_path = root / "drive-scoped.ps1"
        samples_path.write_text(json.dumps(samples), encoding="utf-8")
        driver_path.write_text(SCOPED_DRIVER.format(
            module_dir=str(HERE).replace(chr(92), "/"),
            samples=str(samples_path).replace(chr(92), "/"),
            results=str(results_path).replace(chr(92), "/")), encoding="utf-8")
        result = subprocess.run(["pwsh", "-NoProfile", "-File", str(driver_path)],
                                capture_output=True, text=True, encoding="utf-8",
                                errors="replace", cwd=REPO_ROOT)
        if result.returncode != 0 or not results_path.exists():
            check("the scoped scanner runs", False)
            print(result.stdout + result.stderr)
            return
        counts = json.loads(results_path.read_text(encoding="utf-8"))
        check("the scoped scanner answered for every sample", len(counts) == len(SCOPED))
        if len(counts) != len(SCOPED):
            return
        for (_, description, _, must_refuse), found in zip(SCOPED, counts):
            check(description, (found > 0) == must_refuse)


def test_leak_scanner() -> None:
    """R3-G07: every leak category needs a string it refuses and one it must not."""
    print("R3-G07 leak scanner categories")

    samples = [text for _, text in MUST_REFUSE] + [text for _, text in MUST_ALLOW]
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp).resolve()
        samples_path = root / "samples.json"
        results_path = root / "results.json"
        driver_path = root / "drive-scanner.ps1"

        samples_path.write_text(json.dumps(samples), encoding="utf-8")
        driver_path.write_text(SCANNER_DRIVER.format(
            module_dir=str(HERE).replace(chr(92), "/"),
            samples=str(samples_path).replace(chr(92), "/"),
            results=str(results_path).replace(chr(92), "/")), encoding="utf-8")

        result = subprocess.run(
            ["pwsh", "-NoProfile", "-File", str(driver_path)],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            cwd=REPO_ROOT)
        if result.returncode != 0 or not results_path.exists():
            check("the scanner runs", False)
            print(result.stdout + result.stderr)
            return

        counts = json.loads(results_path.read_text(encoding="utf-8"))
        check("the scanner answered for every sample", len(counts) == len(samples))
        if len(counts) != len(samples):
            return

        for (description, _), found in zip(MUST_REFUSE, counts):
            check(f"{description} is refused", found > 0)
        for (description, _), found in zip(MUST_ALLOW, counts[len(MUST_REFUSE):]):
            check(f"{description} is allowed", found == 0)


def workflow_trigger_paths(workflow: Path) -> set[str]:
    """Reads every glob under a workflow's `paths:` filters.

    Parsed by hand rather than with a YAML library: this fixture runs on a bare CI
    interpreter, and a missing third-party import would turn the check off silently.

    <param name="workflow">Workflow file to read.</param>
    <returns>Every glob listed under any `paths:` key.</returns>
    """
    globs: set[str] = set()
    collecting_at: int | None = None

    for line in workflow.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        indent = len(line) - len(line.lstrip())

        if stripped == "paths:":
            collecting_at = indent
            continue
        if collecting_at is None:
            continue
        if stripped.startswith("#") or not stripped:
            continue
        if stripped.startswith("- ") and indent > collecting_at:
            globs.add(stripped[2:].strip().strip("\'\""))
            continue
        collecting_at = None

    return globs


def matches_glob(path: str, pattern: str) -> bool:
    """Whether a repository-relative path is selected by a workflow path filter."""
    return fnmatch.fnmatchcase(path, pattern.replace("**", "*"))


def test_workflow_triggers() -> None:
    """R3-G06: changing a Graphify script must run its tests and rebuild the page."""
    print("R3-G06 workflow triggers")

    workflows = {
        "the test workflow": REPO_ROOT / ".github" / "workflows" / "test.yml",
        "the Pages workflow": REPO_ROOT / ".github" / "workflows" / "deploy-pages.yml",
    }
    scripts = sorted(f.name for f in HERE.iterdir()
                     if f.is_file() and f.suffix in (".py", ".ps1", ".psm1", ".version"))
    check("the fixture found the Graphify scripts", len(scripts) >= 6)

    for description, workflow in workflows.items():
        globs = workflow_trigger_paths(workflow)
        check(f"{description} declares path filters", bool(globs))
        unmatched = [name for name in scripts
                     if not any(matches_glob(f"graphify/{name}", g) for g in globs)]
        check(f"{description} is triggered by every Graphify script",
              not unmatched)
        if unmatched:
            print(f"        not matched: {unmatched}")

    # A workflow that triggers but does not run the fixtures is no better than one that
    # does not trigger at all.
    for description, workflow in workflows.items():
        body = workflow.read_text(encoding="utf-8")
        check(f"{description} runs the self-test", "graphify/self-test.py" in body)

    pages = workflows["the Pages workflow"].read_text(encoding="utf-8")
    check("the Pages workflow still runs the publish gate",
          "verify-public-map.ps1" in pages)


def run_gate(map_dir: Path, python_on_path: Path | None = None,
             ci: dict[str, str] | None = None, require_commit: str | None = None) -> int:
    """Runs the publish gate against a fixture directory and returns its exit code.

    <param name="map_dir">Directory to verify.</param>
    <param name="python_on_path">Directory to prepend to PATH, to stand in for a broken interpreter.</param>
    <param name="ci">GitHub Actions variables to set; both are cleared otherwise, because the
    fixture map does not describe the checkout a real CI run would be publishing.</param>
    <param name="require_commit">Commit to pin the map to, for the -RequireCommit path.</param>
    <returns>The gate's exit code.</returns>
    """
    env = dict(**__import__("os").environ)
    env.pop("GITHUB_ACTIONS", None)
    env.pop("GITHUB_SHA", None)
    env.update(ci or {})
    if python_on_path is not None:
        env["PATH"] = str(python_on_path) + __import__("os").pathsep + env["PATH"]
    command = ["pwsh", "-NoProfile", "-File", str(HERE / "verify-public-map.ps1"),
               "-MapDirectory", str(map_dir)]
    if require_commit is not None:
        command += ["-RequireCommit", require_commit]

    result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8",
                            errors="replace", cwd=REPO_ROOT, env=env)
    return result.returncode


def republish(map_dir: Path, page: str, metadata: dict) -> None:
    """Writes a page and the metadata that describes it, the way the publisher does.

    <param name="map_dir">The map directory.</param>
    <param name="page">The page to publish.</param>
    <param name="metadata">The metadata to publish alongside it.</param>
    """
    # newline="\n" for the same reason the publisher uses it: these artifacts are pinned to eol=lf,
    # so a fixture writing them the platform way would not be publishing what the publisher does.
    (map_dir / "index.html").write_text(page, encoding="utf-8", newline="\n")

    published = dict(metadata.get("published_sha256") or {})
    published["index.html"] = hashlib.sha256(
        (map_dir / "index.html").read_bytes()).hexdigest()
    metadata = dict(metadata)
    metadata["published_sha256"] = published
    (map_dir / "metadata.json").write_text(json.dumps(metadata, indent=2), encoding="utf-8",
                                           newline="\n")


def fixture_map(docs_root: Path) -> Path:
    """Mirrors the published docs layout into a writable fixture directory.

    The page loads its vendored library through `../assets/`, so the sibling directory has
    to come along or the gate reports a missing asset rather than the case under test.

    <param name="docs_root">Directory to build the docs/ mirror in.</param>
    <returns>Path of the copied architecture-map directory.</returns>
    """
    shutil.copytree(REPO_ROOT / "docs" / "assets", docs_root / "assets")
    destination = docs_root / "architecture-map"
    shutil.copytree(REPO_ROOT / "docs" / "architecture-map", destination)
    return destination


def test_artifacts_derive_from_each_other() -> None:
    """R4-G02: matching ids are not derivation; the content has to agree."""
    print("R4-G02 artifacts derive from each other")
    builder = load("build_public_map_derive", "build-public-map.py")
    bind = builder.ensure_artifacts_share_one_run

    def page(ids, edges=None):
        drawn = [{"id": i, "label": i} for i in ids]
        # A community view has to draw the aggregation of the graph's cross-community edges; an
        # empty array is itself a page that does not draw this graph (R4-G02).
        if edges is None:
            edges = ([{"from": ids[0], "to": ids[1], "label": "1 cross-community edges",
                       "confidence": "AGGREGATED"}] if len(ids) > 1 else [])
        return ("const RAW_NODES = " + json.dumps(drawn) + ";" + chr(10)
                + "const RAW_EDGES = " + json.dumps(edges) + ";")

    extraction = {
        "nodes": [{"id": "a", "label": "A.cs", "source_file": "Core/A.cs"},
                  {"id": "b", "label": "B.cs", "source_file": "Core/B.cs"}],
        "edges": [{"source": "a", "target": "b", "relation": "calls"}],
    }
    graph = {
        "nodes": [{"id": "a", "label": "Core/A.cs", "source_file": "Core/A.cs", "community": "0"},
                  {"id": "b", "label": "B.cs", "source_file": "Core/B.cs", "community": "1"}],
        "links": [{"source": "a", "target": "b", "relation": "calls", "weight": 1.0}],
    }

    allows(lambda: bind(extraction, graph, page(["0", "1"])),
           "a graph and page derived from this extraction are accepted")

    # Every one of these keeps the ids identical, which is all the previous check compared.
    relabelled = json.loads(json.dumps(graph))
    relabelled["nodes"][1]["label"] = "SomethingElse.cs"
    refuses(lambda: bind(extraction, relabelled, page(["0", "1"])),
            "a graph whose labels differ from the extraction is refused")

    moved = json.loads(json.dumps(graph))
    moved["nodes"][1]["source_file"] = "Helpers/B.cs"
    refuses(lambda: bind(extraction, moved, page(["0", "1"])),
            "a graph attributing a node to another file is refused")

    reversed_edge = json.loads(json.dumps(graph))
    reversed_edge["links"] = [{"source": "b", "target": "a", "relation": "calls",
                              "weight": 1.0}]
    refuses(lambda: bind(extraction, reversed_edge, page(["0", "1"])),
            "a graph whose edge runs the other way is refused")

    other_relation = json.loads(json.dumps(graph))
    other_relation["links"] = [{"source": "a", "target": "b", "relation": "implements",
                               "weight": 1.0}]
    refuses(lambda: bind(extraction, other_relation, page(["0", "1"])),
            "a graph giving an edge another relation is refused")

    missing_node = {"nodes": [graph["nodes"][0]], "links": []}
    refuses(lambda: bind(extraction, missing_node, page(["0"], [])),
            "a graph missing a node the extraction has is refused")

    repartitioned = json.loads(json.dumps(graph))
    repartitioned["nodes"][0]["community"] = "7"
    repartitioned["nodes"][1]["community"] = "8"
    refuses(lambda: bind(extraction, repartitioned, page(["0", "1"])),
            "a page drawn from another partition is refused")

    # R8-G04: the checks above compare presence. These five keep every id, every pair and every
    # label the old check looked at, and each one still says something the extraction did not.
    reweighted = json.loads(json.dumps(graph))
    reweighted["links"][0]["weight"] = 999
    refuses(lambda: bind(extraction, reweighted, page(["0", "1"])),
            "a graph that reweights a relationship is refused")

    unweighted = json.loads(json.dumps(graph))
    unweighted["links"][0].pop("weight")
    refuses(lambda: bind(extraction, unweighted, page(["0", "1"])),
            "a graph that carries no weight at all is refused")

    # Two relations for one pair. Dropping the `calls` edge leaves the pair drawn, so a check that
    # asked only whether the pair appeared saw nothing wrong.
    two_relations = json.loads(json.dumps(extraction))
    two_relations["edges"].append({"source": "a", "target": "b", "relation": "references"})
    both_drawn = json.loads(json.dumps(graph))
    both_drawn["links"].append({"source": "a", "target": "b", "relation": "references",
                                "weight": 1.0})
    two_edge_page = page(["0", "1"], [{"from": "0", "to": "1",
                                       "label": "2 cross-community edges",
                                       "confidence": "AGGREGATED"}])
    allows(lambda: bind(two_relations, both_drawn, two_edge_page),
           "a graph drawing both relations for a pair is accepted")

    # The one collapse this build really performs, measured over the published artifacts: a
    # `references` edge disappears when the pair is drawn under a stronger relation.
    collapsed = json.loads(json.dumps(graph))
    allows(lambda: bind(two_relations, collapsed, page(["0", "1"])),
           "a references edge absorbed into a stronger relation is still accepted")

    dropped_calls = json.loads(json.dumps(graph))
    dropped_calls["links"] = [{"source": "a", "target": "b", "relation": "references"}]
    refuses(lambda: bind(two_relations, dropped_calls, page(["0", "1"])),
            "a graph that drops one relation while keeping the pair is refused")

    # A label may carry a path prefix, and nothing else.
    prefixed = json.loads(json.dumps(graph))
    prefixed["nodes"][1]["label"] = "Deprecated: B.cs"
    refuses(lambda: bind(extraction, prefixed, page(["0", "1"])),
            "a label given an arbitrary prefix is refused")

    path_prefixed = json.loads(json.dumps(graph))
    path_prefixed["nodes"][1]["label"] = "Core/B.cs"
    allows(lambda: bind(extraction, path_prefixed, page(["0", "1"])),
           "a label carrying the path prefix this build adds is accepted")

    # Removing the attribution used to skip the comparison entirely.
    unattributed = json.loads(json.dumps(graph))
    del unattributed["nodes"][1]["source_file"]
    refuses(lambda: bind(extraction, unattributed, page(["0", "1"])),
            "a graph that drops an attribution the extraction has is refused")

    refuses(lambda: bind(extraction, graph, page(["0"])),
            "a page drawing only some of the communities is refused")


def test_lossless_relationship_artifact() -> None:
    """R31-G01: the public artifacts retain relationships a DiGraph projection collapses."""
    print("R31-G01 lossless relationship artifact")
    builder = load("build_public_map_relationships", "build-public-map.py")
    if not hasattr(builder, "relationship_catalog"):
        check("the public builder emits a lossless relationship catalog", False)
        return

    extraction = {
        "nodes": [
            {"id": "a", "label": "A.cs", "source_file": "Core/A.cs"},
            {"id": "b", "label": "B.cs", "source_file": "Core/B.cs"},
        ],
        "edges": [
            {"source": "a", "target": "b", "relation": "calls",
             "confidence": "EXTRACTED", "source_file": "Core/A.cs"},
            {"source": "a", "target": "b", "relation": "references",
             "confidence": "EXTRACTED", "source_file": "Core/A.cs"},
        ],
    }
    catalog = builder.relationship_catalog(extraction, Path("."), "fixture-run")
    relationships = builder.catalog_relationships(catalog)
    check("both same-endpoint relations survive in the lossless catalog",
          catalog.get("schema") == 1
          and catalog.get("format") == "grouped-fields"
          and catalog.get("corpus_run_id") == "fixture-run"
          and catalog.get("relationship_count") == 2
          and {edge.get("relation") for edge in relationships} == {"calls", "references"})

    altered = json.loads(json.dumps(catalog))
    altered["groups"][0]["rows"].pop()
    refuses(lambda: builder.ensure_relationship_catalog(extraction, altered, Path("."),
                                                          "fixture-run"),
            "a catalog that drops one collapsed relationship is refused")


def test_deleted_sources_count_as_dirty() -> None:
    """R4-G03: a corpus file removed from the tree is still a corpus path."""
    print("R4-G03 deleted sources are recognised")
    allowlist = load("corpus_allowlist_membership", "corpus_allowlist.py")

    for path in ("Helpers/Deleted.cs", "Core/Session/Gone.cs", "README.md", "docs/faq.html"):
        check(f"{path} is recognised as a corpus path without existing",
              allowlist.belongs_to_corpus(path))

    for path in ("Tests/Foo.cs", "graphify/self-test.py", "docs/architecture-map/index.html",
                 "Helpers.cs", "docs/unlisted.html"):
        check(f"{path} is not a corpus path", not allowlist.belongs_to_corpus(path))


def test_gate_fails_closed() -> None:
    """R2-G02: a provenance helper that cannot run must fail the gate, not be skipped."""
    print("R2-G02 gate fails closed")
    import os

    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp).resolve()
        good = fixture_map(root / "docs")
        check("an unmodified map passes", run_gate(good) == 0)

        # A python on PATH that always fails stands in for a broken CI interpreter.
        stub_dir = root / "stub"
        stub_dir.mkdir()
        if os.name == "nt":
            (stub_dir / "python.cmd").write_text("@exit /b 3\r\n", encoding="utf-8")
        else:
            stub = stub_dir / "python"
            stub.write_text("#!/bin/sh\nexit 3\n", encoding="utf-8")
            stub.chmod(0o755)
        check("a failing corpus-digest helper fails the gate",
              run_gate(good, python_on_path=stub_dir) != 0)

        metadata_path = good / "metadata.json"
        original = json.loads(metadata_path.read_text(encoding="utf-8"))

        for description, mutate in (
            ("a map with no source_files fails the gate", lambda m: m.update(source_files=[])),
            ("a map missing source_files fails the gate", lambda m: m.pop("source_files", None)),
            ("a source outside the corpus fails the gate",
             lambda m: m["source_files"].append("Tests/Secret.cs")),
            ("an overstated rendered node count fails the gate",
             lambda m: m.update(rendered_graph_nodes=999999)),
            ("a map with no corpus run id fails the gate",
             lambda m: m.pop("corpus_run_id", None)),
            ("a map with no artifact digests fails the gate",
             lambda m: m.pop("artifact_sha256", None)),
            ("a map with an unusable artifact digest fails the gate",
             lambda m: m["artifact_sha256"].update({"graph.json": "not-a-digest"})),
            ("a map with no provenance chain fails the gate",
             lambda m: m.pop("provenance_chain", None)),
            ("a provenance chain missing a link fails the gate",
             lambda m: m["provenance_chain"].pop("graph_link", None)),
            ("a provenance chain whose digest disagrees with artifact_sha256 fails the gate",
             lambda m: m["provenance_chain"].update({"graph_sha256": "b" * 64})),
            ("a map with no extraction diagnostics fails the gate",
             lambda m: m.pop("diagnostics", None)),
            ("a map with a dangling extraction edge fails the gate",
             lambda m: m["diagnostics"].update({"dangling_endpoint_edges": 1})),
        ):
            mutated = json.loads(json.dumps(original))
            mutate(mutated)
            metadata_path.write_text(json.dumps(mutated, indent=2), encoding="utf-8")
            check(description, run_gate(good) != 0)

        # R3-G05: the source pair, the view name and the banner used to be four independent
        # copies of the same facts, so changing any one of them left the gate silent.
        for description, mutate in (
            ("an inflated source node count fails the gate",
             lambda m: m.update(source_graph_nodes=m["source_graph_nodes"] + 1)),
            ("an inflated source edge count fails the gate",
             lambda m: m.update(source_graph_edges=m["source_graph_edges"] + 1)),
            ("a mislabelled rendered view fails the gate",
             lambda m: m.update(rendered_view="full")),
            ("an understated inferred-edge count fails the gate",
             lambda m: m["edge_confidence"].update(INFERRED=0)),
        ):
            mutated = json.loads(json.dumps(original))
            mutate(mutated)
            metadata_path.write_text(json.dumps(mutated, indent=2), encoding="utf-8")
            check(description, run_gate(good) != 0)

        metadata_path.write_text(json.dumps(original, indent=2), encoding="utf-8")

        # R4-G04: the workflow uploads the whole docs/ tree, but the content scan only looked at
        # five extensions, so anything else here went live unscanned.
        stray = good / "notes.txt"
        stray.write_text("internal note", encoding="utf-8")
        check("an unexpected file in the map directory fails the gate", run_gate(good) != 0)
        stray.unlink()

        unscanned = good / "diagram.svg"
        unscanned.write_text("<svg/>", encoding="utf-8")
        check("a file with an unscanned extension fails the gate", run_gate(good) != 0)
        unscanned.unlink()

        check("the map passes again once the directory holds only its own output",
              run_gate(good) == 0)

        # Read and restore as bytes. Text mode translates line endings both ways, so putting the
        # page "back" through it rewrote an LF page as CRLF and the restore no longer matched the
        # digest it was supposed to restore — a fixture about published bytes cannot go through a
        # layer that edits them.
        index_path = good / "index.html"
        page = index_path.read_bytes()

        # .gitattributes pins the published artifacts to eol=lf. Written the platform way on
        # Windows they carried CRLF, so the digest recorded here described bytes no clean checkout
        # would ever produce: the gate passed on the machine that built the map and would have
        # failed in CI on the same commit (found in the seventh round).
        check("the recorded digest survives a checkout",
              hashlib.sha256(page.replace(b"\r\n", b"\n")).hexdigest()
              == original["published_sha256"]["index.html"])

        edited = page.replace(f"{original['source_graph_nodes']:,} nodes".encode(),
                              f"{original['source_graph_nodes'] + 7:,} nodes".encode(), 1)
        check("the banner fixture actually changed the page", edited != page)
        index_path.write_bytes(edited)
        check("a banner that disagrees with the metadata fails the gate", run_gate(good) != 0)
        index_path.write_bytes(page)

        # R7-G03: the metadata described graphify-out/graph.html, which is the page before the
        # title change, the CDN rewrite, the hardening and the banner. Nothing described the file
        # a reader downloads, so a byte changed afterwards contradicted nothing.
        index_path.write_bytes(page.replace(b"<title>", b"<title> ", 1))
        check("a single byte changed in the published page fails the gate", run_gate(good) != 0)
        index_path.write_bytes(page)
        check("the map passes again once the published bytes are back", run_gate(good) == 0)

        relationships_path = good / "relationships.json"
        relationship_bytes = relationships_path.read_bytes()
        relationships_path.write_bytes(relationship_bytes + b" ")
        check("a single byte changed in the lossless relationship artifact fails the gate",
              run_gate(good) != 0)
        relationships_path.write_bytes(relationship_bytes)
        check("the map passes again once the relationship bytes are back", run_gate(good) == 0)

        # R3-G04: the recorded commit had only to be non-empty, so an uncommitted build
        # published as a description of the current source, and the banner hid the fact.
        #
        # The published map may itself have been built from a dirty tree, which is legitimate
        # locally, so both states are constructed here rather than assumed. A fixture that
        # depends on how the real artifact happened to be built tests the artifact, not the gate.
        recorded = str(original["source_commit"])
        clean_commit = recorded[: -len("-dirty")] if recorded.endswith("-dirty") else recorded

        clean = json.loads(json.dumps(original))
        clean["source_commit"] = clean_commit

        index_path = good / "index.html"
        published_page = index_path.read_text(encoding="utf-8")
        clean_page = published_page.replace("-dirty (uncommitted working tree)", "")
        republish(good, clean_page, clean)

        check("a map still passes with the CI variables set",
              run_gate(good, ci={"GITHUB_ACTIONS": "true", "GITHUB_SHA": "0" * 40}) == 0)
        check("pinning the commit it was built from passes",
              run_gate(good, require_commit=clean_commit) == 0)
        check("pinning a different commit fails the gate",
              run_gate(good, require_commit="0" * 40) != 0)

        # The metadata now says the sources were uncommitted while the page does not disclose it.
        dirty = json.loads(json.dumps(clean))
        dirty["source_commit"] = clean_commit + "-dirty"
        republish(good, clean_page, dirty)

        # The page still shows the clean commit, so it is now hiding the state the metadata
        # records.
        check("a banner that hides the uncommitted state fails the gate", run_gate(good) != 0)
        check("an uncommitted build cannot be pinned to a commit",
              run_gate(good, require_commit=clean_commit) != 0)

        # A map generated from uncommitted sources is publishable — it is always generated
        # before the commit that carries it, and the corpus digest is what proves it describes
        # this source. What it may not do is hide that state from the reader.
        disclosed = clean_page.replace(clean_commit[:12],
                                       clean_commit[:12] + "-dirty (uncommitted working tree)", 1)
        check("the disclosure fixture actually changed the page", disclosed != clean_page)
        republish(good, disclosed, dirty)
        check("an uncommitted build that discloses itself is publishable in CI",
              run_gate(good, ci={"GITHUB_ACTIONS": "true", "GITHUB_SHA": clean_commit}) == 0)

        # R23-G01 at the gate. A decoy `const LEGEND = [];` inside a RAW_NODES label, and the
        # real LEGEND carrying a raw end-tag. The old gate checked the decoy and passed.
        builder = load("build_public_map_gate", "build-public-map.py")
        ns, ne = builder.script_literal(published_page, "RAW_NODES")
        nodes = json.loads(published_page[ns:ne])
        nodes[0]["label"] = "decoy const LEGEND = []; end"
        shadowed = published_page[:ns] + json.dumps(nodes) + published_page[ne:]
        ls, le = builder.script_literal(shadowed, "LEGEND")
        legend = json.loads(shadowed[ls:le])
        legend[0]["label"] = "real </script><script>alert(1)</script>"
        shadowed = shadowed[:ls] + json.dumps(legend) + shadowed[le:]
        republish(good, shadowed, original)
        check("a decoy declaration inside an earlier literal does not hide a raw payload from the gate",
              run_gate(good) != 0)

        # R23-G02 at the gate. A hidden file in the published tree, with a secret in it: the
        # inventory must see it (it is not in the manifest) and so must the scan.
        republish(good, published_page, original)
        hidden = good / ".env"
        hidden.write_text("GITHUB_TOKEN=github_pat_" + "A1b2C3d4E5" * 8 + "Zz" + NL, encoding="utf-8")
        if os.name == "nt":
            os.system(f'attrib +h "{hidden}"')
        check("a hidden file with a secret in the map directory fails the map gate", run_gate(good) != 0)
        hidden.unlink()

        # R22-G01 at the gate. A LEGEND label that carries `]; </script>` raw, with the count
        # unchanged so nothing but the script-context check can object. The old regex ended its
        # match inside the string, saw no `<` in what it matched, and passed the page.
        builder = load("build_public_map_gate", "build-public-map.py")
        start, end = builder.script_literal(published_page, "LEGEND")
        legend = json.loads(published_page[start:end])
        check("the published page has a legend row to poison", len(legend) > 0)
        legend[0]["label"] = "poisoned ]; </script><script>alert(1)</script>"
        poisoned = published_page[:start] + json.dumps(legend) + published_page[end:]
        republish(good, poisoned, original)
        check("a raw terminator payload inside a LEGEND label fails the gate", run_gate(good) != 0)

        # And the map goes back to exactly what was published, metadata included.
        republish(good, published_page, original)
        check("the restored map passes again", run_gate(good) == 0)


def test_page_hardening() -> None:
    """R2-G05: graph text must not be able to leave its script or HTML context."""
    print("R2-G05 published page hardening")
    builder = load("build_public_map", "build-public-map.py")

    label = "</script><img src=x onerror=alert(1)>"
    page = ('const RAW_NODES = [{"label": ' + json.dumps(label) + '}];\n'
            'const RAW_EDGES = [];\n'
            'const LEGEND = [{"label": ' + json.dumps(label) + ', "color": "#fff"}];\n'
            'const hyperedges = [{"id": "h1", "label": ' + json.dumps(label) + '}];\n'
            '${c.label} ${c.color} ${c.count}\n')
    hardened = builder.harden_page(page)

    check("the script element can no longer be closed early", "</script>" not in hardened)
    check("the legend literal is hardened too (R20-G02)",
          "<" not in hardened.split("const LEGEND")[1].split(";")[0])
    # R21-G03's widened guard found this one on the real page: lower-case, source-derived,
    # emitted by the template on every page, never hardened.
    check("the hyperedges literal is hardened too (R21-G03)",
          "<" not in hardened.split("const hyperedges")[1].split(";")[0])
    refuses(lambda: builder.harden_page(page.replace("const hyperedges", "const hyperedgesX")),
            "a page without the hyperedges literal is refused, not silently passed")

    stray = page + 'const HYPEREDGES = [{"label": ' + json.dumps(label) + '}];\n'
    refuses(lambda: builder.harden_page(stray),
            "an inline constant the hardening does not know is refused")
    # R21-G03: the guard used to see ALL_CAPS names only. It judges the right-hand side: a
    # non-empty array, object or string literal is data the page would write unescaped.
    for name in ("legendData", "raw_nodes2", "_extra", "$legend"):
        refuses(lambda n=name: builder.harden_page(page + 'const ' + n + ' = [{"x": 1}];\n'),
                f"an inline constant named {name} is refused too")
    refuses(lambda: builder.harden_page(page + 'const title = "</script>";\n'),
            "an inline string constant is refused")
    refuses(lambda: builder.harden_page(page + 'const meta = { "label": "x" };\n'),
            "an inline object constant is refused")
    allows(lambda: builder.harden_page(page + "const el = document.createElement('div');\n"),
           "a constant assigned from an expression is code, not a data literal")
    allows(lambda: builder.harden_page(page + "const out = [];\nconst seen = {};\n"),
           "an exactly empty array or object carries nothing")
    check("no raw angle bracket survives in the data array",
          "<" not in hardened.split("const RAW_EDGES")[0])
    check("the legend escapes its label", "${esc(c.label)}" in hardened)
    check("the legend escapes its colour", "${esc(c.color)}" in hardened)

    restored = json.loads(hardened.split("const RAW_NODES = ")[1].split(";")[0])
    check("the escaped label still decodes to the original text",
          restored[0]["label"] == label)

    # R22-G01. A label may legally contain `];` or `};`; the regex that used to find the
    # literal ended there, escaped the prefix and left the rest raw — and the verifier, using
    # the same regex, checked the same prefix. Every literal, every terminator shape.
    terminators = {
        "an array terminator": "prefix ]; </script><script>alert(1)</script>",
        "an object terminator": "prefix }; </script><script>alert(1)</script>",
        "an escaped quote": 'quote \" then ]; </script>',
        "a backslash": "back\\slash ]; </script>",
        "a line separator": "line\u2028break ]; </script>",
    }
    nested = {"a": [1, {"b": "]; </script>"}], "c": {"d": ["}; </script>"]}}
    for what, text in terminators.items():
        for name in ("RAW_NODES", "LEGEND", "hyperedges"):
            row = {"id": "n", "label": text, "color": "#fff", "nested": nested}
            tricky = page.replace(
                "const " + name + " = " + hardened_source(page, name) + ";",
                "const " + name + " = " + json.dumps([row]) + ";", 1)
            check(f"{name} fixture with {what} was planted", tricky != page)
            try:
                out = builder.harden_page(tricky)
            except SystemExit as refusal:
                check(f"{name} with {what}: the hardener handles the payload rather than aborting ({refusal})", False)
                continue
            start, end = builder.script_literal(out, name)
            literal = out[start:end]
            check(f"{name} with {what}: no raw < survives anywhere in the literal",
                  "<" not in literal and ">" not in literal)
            check(f"{name} with {what}: no raw line or paragraph separator survives",
                  chr(0x2028) not in literal and chr(0x2029) not in literal)
            check(f"{name} with {what}: the literal still decodes to the original data",
                  json.loads(literal) == [row])
            check(f"{name} with {what}: nothing after the literal was mistaken for it",
                  "</script><script>" not in out.split("const " + name)[1].split(";", 1)[0])
    refuses(lambda: builder.harden_page(page.replace('const RAW_EDGES = [];', 'const RAW_EDGES = [;')),
            "an unbalanced literal is refused, not hardened as far as it goes")

    # R23-G01. A decoy declaration inside an earlier literal's string used to be the one found,
    # so the real, later literal stayed raw. Every ordered pair of literals.
    names = ("RAW_NODES", "RAW_EDGES", "LEGEND", "hyperedges")
    for host in names:
        for target in names:
            if host == target:
                continue
            decoy = json.dumps([{"id": "d", "label": "decoy const " + target + " = []; end", "color": "#fff"}])
            payload = json.dumps([{"id": "p", "label": "</script><script>alert(1)</script>", "color": "#fff"}])
            # Declarations are replaced where they start a line — the very rule under test; the
            # first "const NAME =" in the text is exactly what a decoy is designed to be.
            shadowed = replace_declaration(page, host, decoy)
            shadowed = replace_declaration(shadowed, target, payload)
            check(f"{host} shadowing {target}: fixture planted", shadowed.count("const " + target + " = ") == 2)
            try:
                out = builder.harden_page(shadowed)
                check(f"{host} shadowing {target}: the real literal is the one hardened",
                      "</script>" not in out and "alert(1)" in out)
            except SystemExit as refusal:
                check(f"{host} shadowing {target}: the hardener handles the page rather than aborting ({refusal})", False)
            # The decoy at the start of a line — only a raw line break inside the string can put
            # it there, which JSON never emits — is two declarations, and refused.
            at_line_start = shadowed.replace("decoy const " + target, "decoy" + NL + "const " + target, 1)
            refuses(lambda s=at_line_start: builder.harden_page(s),
                    f"{host} shadowing {target}: a second declaration at a line start is refused")
    refuses(lambda: builder.harden_page(page.replace('const RAW_EDGES = [];', 'const RAW_EDGES = []')),
            "a literal without its semicolon is refused")


def replace_declaration(page: str, name: str, literal: str) -> str:
    """Replaces the statement `const NAME = ...;` that starts a line, and only that one.

    <param name="page">The fixture page.</param>
    <param name="name">The constant's name.</param>
    <param name="literal">The new literal text.</param>
    <returns>The page with that one declaration replaced.</returns>
    """
    import re
    pattern = re.compile(r"(?m)^const " + re.escape(name) + r" = .*?;$")
    assert len(pattern.findall(page)) == 1, name
    return pattern.sub(lambda _: "const " + name + " = " + literal + ";", page, count=1)


def hardened_source(page: str, name: str) -> str:
    """The literal text a fixture page currently carries for a name.

    <param name="page">The fixture page.</param>
    <param name="name">The constant's name.</param>
    <returns>The literal, exclusive of its semicolon.</returns>
    """
    head = page.index("const " + name + " = ") + len("const " + name + " = ")
    return page[head:page.index(";", head)]


def test_work_directory_is_accounted_for() -> None:
    """R4-G01: an inventory proves what the tooling produced; it never launders what it finds."""
    print("R4-G01 the work inventory only accounts for what it can prove")
    allowlist = load("corpus_allowlist_work", "corpus_allowlist.py")
    stage = load("stage_corpus_work", "stage-corpus.py")
    run = "0123456789abcdef0123456789abcdef"

    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        directory = stage_a_directory(stage, root, "staged")
        work = directory / allowlist.STAGING_WORK_DIR
        work.mkdir(parents=True, exist_ok=True)
        (work / "graph.json").write_text("{}", encoding="utf-8")
        (work / "smuggled.txt").write_text("someone else's file", encoding="utf-8")
        recorded = {"Program.cs"}

        def unlisted():
            return allowlist.unlisted_files(directory, recorded, run_id=run)

        check("an extraction output nothing has recorded is unaccounted for",
              "graphify-out/graph.json" in unlisted())
        refuses(lambda: stage.ensure_safe_staging_target(directory),
                "--force is refused while the work directory is unaccounted for")

        # What a successful publish records: the artifacts it actually consumed, and nothing else.
        check("the inventory records only what the caller named",
              allowlist.write_work_inventory(
                  directory, {"graphify-out/graph.json": "graph"}, run) == 1)
        check("the recorded artifact is accounted for",
              "graphify-out/graph.json" not in unlisted())

        # The counter-example: a refresh used to walk the directory, so an unknown file planted
        # beforehand was laundered into the inventory and --force would then delete it (R4-G01).
        check("a file the refresh never named stays unknown",
              unlisted() == ["graphify-out/smuggled.txt"])
        refuses(lambda: stage.ensure_safe_staging_target(directory),
                "--force stays refused for the file nobody recorded")

        # An entry only counts while the bytes are the ones that were recorded.
        (work / "graph.json").write_text('{"edited": true}', encoding="utf-8")
        check("a recorded artifact whose bytes changed stops being accounted for",
              "graphify-out/graph.json" in unlisted())

        allowlist.write_work_inventory(directory, {"graphify-out/graph.json": "graph"}, run)
        check("re-recording it accounts for it again",
              "graphify-out/graph.json" not in unlisted())
        check("an inventory from another staging run accounts for nothing",
              "graphify-out/graph.json" in
              allowlist.unlisted_files(directory, recorded, run_id="f" * 32))

        # An unreadable or wrong-schema inventory is not a licence to delete either.
        allowlist.work_inventory_path(directory).write_text("not json", encoding="utf-8")
        check("an unreadable inventory accounts for nothing rather than exempting everything",
              unlisted() == ["graphify-out/graph.json", "graphify-out/smuggled.txt"])
        allowlist.work_inventory_path(directory).write_text(
            json.dumps({"schema": 1, "run_id": run,
                        "entries": [{"path": "graphify-out/graph.json", "sha256": "0" * 64}]}),
            encoding="utf-8")
        check("an inventory written to an older schema accounts for nothing",
              "graphify-out/graph.json" in unlisted())


def test_page_payload_matches() -> None:
    """R4-G02: the page's content and its whole edge set have to come from this graph."""
    print("R4-G02 the page draws this graph's edges, and only them")
    builder = load("build_public_map_payload", "build-public-map.py")
    bind = builder.ensure_artifacts_share_one_run

    def page(nodes, edges):
        return ("const RAW_NODES = " + json.dumps(nodes) + ";" + chr(10)
                + "const RAW_EDGES = " + json.dumps(edges) + ";")

    def edge(source, target, label, confidence="EXTRACTED"):
        return {"from": source, "to": target, "label": label, "confidence": confidence}

    extraction = {
        "nodes": [{"id": "a", "label": "Alpha"}, {"id": "b", "label": "Beta"}],
        "edges": [{"source": "a", "target": "b", "relation": "calls"}],
    }
    graph = {
        "nodes": [{"id": "a", "label": "Alpha", "community": "0", "community_name": "Core"},
                  {"id": "b", "label": "Beta", "community": "1", "community_name": "Edge"}],
        "edges": [{"source": "a", "target": "b", "relation": "calls",
                   "confidence": "EXTRACTED", "weight": 1.0}],
    }
    node_view = [{"id": "a", "label": "Alpha"}, {"id": "b", "label": "Beta"}]
    agg_view = [{"id": "0", "label": "Core"}, {"id": "1", "label": "Edge"}]

    allows(lambda: bind(extraction, graph, page(node_view, [edge("a", "b", "calls")])),
           "a node view drawing this graph's own edge is accepted")
    allows(lambda: bind(extraction, graph,
                        page(agg_view, [edge("0", "1", "1 cross-community edges", "AGGREGATED")])),
           "a community view carrying the graph's own aggregation is accepted")

    # Only the endpoints were checked before, so every one of these passed (R4-G02).
    for description, nodes, edges in (
        ("an edge drawn the wrong way round", node_view, [edge("b", "a", "calls")]),
        ("an edge whose relation was changed", node_view, [edge("a", "b", "implements")]),
        ("an edge whose confidence was changed",
         node_view, [edge("a", "b", "calls", "INFERRED")]),
        ("an edge drawn twice", node_view, [edge("a", "b", "calls"), edge("a", "b", "calls")]),
        ("a node view that draws no edges at all", node_view, []),
        ("an aggregated count that was inflated",
         agg_view, [edge("0", "1", "9 cross-community edges", "AGGREGATED")]),
        ("an aggregated pair the graph does not bridge",
         agg_view, [edge("0", "1", "1 cross-community edges", "AGGREGATED"),
                    edge("1", "0", "1 cross-community edges", "AGGREGATED")]),
        ("a community view that draws no edges at all", agg_view, []),
        ("an aggregated edge that hides how many it stands for",
         agg_view, [edge("0", "1", "cross-community edges", "AGGREGATED")]),
    ):
        refuses(lambda n=nodes, e=edges: bind(extraction, graph, page(n, e)),
                f"{description} is refused")

    # R7-G03: the three the sixth round's comparison still let through.
    confidence_rewritten = json.loads(json.dumps(graph))
    confidence_rewritten["edges"][0]["confidence"] = "INFERRED"
    refuses(lambda: bind(extraction, confidence_rewritten,
                         page(agg_view, [edge("0", "1", "1 cross-community edges", "AGGREGATED")])),
            "a graph that relabels an EXTRACTED edge as INFERRED is refused")

    refuses(lambda: bind(extraction, graph,
                         page([{"id": "a"}, {"id": "b", "label": "Beta"}],
                              [edge("a", "b", "calls")])),
            "a page node drawn with no label at all is refused")

    refuses(lambda: bind(extraction, graph,
                         page(agg_view,
                              [edge("0", "1", "1 cross-community edges", "AGGREGATED"),
                               edge("0", "1", "0 cross-community edges", "AGGREGATED")])),
            "an aggregated count split across two entries for the same pair is refused")

    refuses(lambda: bind(extraction, graph,
                         page(agg_view,
                              [edge("0", "1", "1 cross-community edges and anything else",
                                    "AGGREGATED")])),
            "an aggregated label carrying anything beyond its count is refused")

    # The label and node-set checks the fourth round added still hold.
    refuses(lambda: bind(extraction, graph,
                         page([{"id": "0", "label": "FORGED"}],
                              [edge("0", "1", "1 cross-community edges", "AGGREGATED")])),
            "a page that renames the graph's communities is refused")
    refuses(lambda: bind(extraction, {"nodes": graph["nodes"], "edges": []},
                         page(agg_view, [])),
            "a graph that drops the extraction's relationships is refused")


def test_artifact_forgeries_refused() -> None:
    """R9-G01: the equivalence check accepted six forged artifacts; none may pass now.

    Each was measured against the live checker before the rules were tightened, and each was
    accepted: a label given an arbitrary prefix, an attribution shortened to a substring of the
    real one, a weight deleted rather than falsified, an edge duplicated with the page's
    aggregate count moved to match, and a dropped `references` edge covered by something other
    than the strong relation that actually absorbs it.
    """
    print("R9-G01 forged artifacts are refused")
    builder = load("build_public_map_forgery", "build-public-map.py")
    bind = builder.ensure_artifacts_share_one_run

    def page(nodes, edges):
        drawn = [{"id": n["id"], "label": n.get("label")} for n in nodes]
        raw = [{"from": e["source"], "to": e["target"], "label": e.get("relation", ""),
                "confidence": e.get("confidence", "EXTRACTED")} for e in edges]
        return ("const RAW_NODES = " + json.dumps(drawn) + ";" + chr(10)
                + "const RAW_EDGES = " + json.dumps(raw) + ";")

    def community_page(ids, counts):
        drawn = [{"id": i} for i in ids]
        raw = [{"from": a, "to": b, "label": f"{n} cross-community edges"}
               for (a, b), n in counts.items()]
        return ("const RAW_NODES = " + json.dumps(drawn) + ";" + chr(10)
                + "const RAW_EDGES = " + json.dumps(raw) + ";")

    def honest():
        extraction = {
            "nodes": [{"id": "n1", "label": "B.cs", "source_file": "Core/B.cs"},
                      {"id": "n2", "label": "C.cs", "source_file": "Core/C.cs"}],
            "edges": [{"source": "n1", "target": "n2", "relation": "calls",
                       "confidence": "EXTRACTED"}],
        }
        graph = {
            "nodes": [{"id": "n1", "label": "B.cs", "source_file": "Core/B.cs", "community": "0"},
                      {"id": "n2", "label": "C.cs", "source_file": "Core/C.cs", "community": "1"}],
            "links": [{"source": "n1", "target": "n2", "relation": "calls",
                       "confidence": "EXTRACTED", "weight": 1.0}],
        }
        return extraction, graph

    extraction, graph = honest()
    allows(lambda: bind(extraction, graph, page(graph["nodes"], graph["links"])),
           "artifacts that agree with each other are accepted")

    # 1. The prefix has to come from that node's own path, not be any text ending in a slash.
    extraction, forged = honest()
    forged["nodes"][0]["label"] = "Deprecated/B.cs"
    refuses(lambda: bind(extraction, forged, page(forged["nodes"], forged["links"])),
            "a label given a prefix that is not from its own source path is refused")

    # 2. Attribution was compared with `in`, so the real path's tail was accepted for the path.
    extraction, forged = honest()
    forged["nodes"][0]["source_file"] = "B.cs"
    refuses(lambda: bind(extraction, forged, page(forged["nodes"], forged["links"])),
            "an attribution shortened to a substring of the extraction's is refused")

    # 3. Checking weight only when present made deleting it a way of skipping the check.
    extraction, forged = honest()
    forged["links"][0].pop("weight")
    refuses(lambda: bind(extraction, forged, page(forged["nodes"], forged["links"])),
            "an edge that carries no weight at all is refused")

    # 4. Edges were compared as a set, so one extraction edge could be drawn twice and the page's
    #    aggregate count raised to agree with the inflated graph.
    extraction, forged = honest()
    forged["links"].append(dict(forged["links"][0]))
    refuses(lambda: bind(extraction, forged,
                         community_page(["0", "1"], {("0", "1"): 2})),
            "an edge duplicated, with the page's count raised to match, is refused")

    # 5/6. A dropped `references` edge is only a collapse when a strong relation absorbed it.
    for description, remaining in (
        ("references [EXTRACTED]", {"relation": "references", "confidence": "EXTRACTED"}),
        ("a weak 'mentions' relation", {"relation": "mentions", "confidence": "INFERRED"}),
    ):
        _, forged = honest()
        offered = {
            "nodes": [{"id": "n1", "label": "B.cs", "source_file": "Core/B.cs"},
                      {"id": "n2", "label": "C.cs", "source_file": "Core/C.cs"}],
            "edges": [{"source": "n1", "target": "n2", "relation": "references",
                       "confidence": "INFERRED"},
                      dict({"source": "n1", "target": "n2"}, **remaining)],
        }
        forged["links"] = [dict({"source": "n1", "target": "n2", "weight": 1.0}, **remaining)]
        refuses(lambda o=offered, f=forged: bind(o, f, page(f["nodes"], f["links"])),
                f"a references edge left covered only by {description} is refused")

    # The collapse the build really performs stays accepted, for each strong relation.
    for strong in ("calls", "defines", "implements", "inherits"):
        _, collapsed = honest()
        offered = {
            "nodes": [{"id": "n1", "label": "B.cs", "source_file": "Core/B.cs"},
                      {"id": "n2", "label": "C.cs", "source_file": "Core/C.cs"}],
            "edges": [{"source": "n1", "target": "n2", "relation": "references",
                       "confidence": "INFERRED"},
                      {"source": "n1", "target": "n2", "relation": strong,
                       "confidence": "EXTRACTED"}],
        }
        collapsed["links"] = [{"source": "n1", "target": "n2", "relation": strong,
                               "confidence": "EXTRACTED", "weight": 1.0}]
        allows(lambda o=offered, c=collapsed: bind(o, c, page(c["nodes"], c["links"])),
               f"a references edge collapsed into {strong} is accepted")

    # The prefix the build really adds — a run of segments from the node's own path — stays
    # accepted, or 18 of the real map's labels would be refused.
    prefixed_extraction = {
        "nodes": [{"id": "n1", "label": "PropertiesTool.cs",
                   "source_file": "Tools/Excel/Properties/PropertiesTool.cs"}],
        "edges": [],
    }
    prefixed_graph = {
        "nodes": [{"id": "n1", "label": "Excel/Properties/PropertiesTool.cs",
                   "source_file": "Tools/Excel/Properties/PropertiesTool.cs", "community": "0"}],
        "links": [],
    }
    allows(lambda: bind(prefixed_extraction, prefixed_graph, page(prefixed_graph["nodes"], [])),
           "a label prefixed with segments of its own source path is accepted")

    # §21.5: a contiguous run of segments was still wider than the builder. The prefix it emits
    # is always a suffix of the parent path, so a leading segment and the basename are not it.
    for bad_prefix, description in (("Tools", "a leading segment of the path"),
                                    ("PropertiesTool.cs", "the file's own name"),
                                    ("Tools/PropertiesTool.cs", "a run that includes the filename")):
        forged_prefix = json.loads(json.dumps(prefixed_graph))
        forged_prefix["nodes"][0]["label"] = bad_prefix + "/PropertiesTool.cs"
        refuses(lambda e=prefixed_extraction, g=forged_prefix:
                bind(e, g, page(g["nodes"], [])),
                f"a label prefixed with {description} is refused")

    # The extraction records where it read a file; the graph records the repository path. Those
    # are the same attribution, and 189 of the real map's nodes depend on it.
    staging_root = (Path(tempfile.gettempdir()) / "graphify-staging-run").resolve()
    staged = {
        "nodes": [{"id": "n1", "label": "developers.html",
                   "source_file": str(staging_root / "docs" / "developers.html")}],
        "edges": [],
    }
    relative = {
        "nodes": [{"id": "n1", "label": "developers.html",
                   "source_file": "docs/developers.html", "community": "0"}],
        "links": [],
    }
    allows(lambda: bind(staged, relative, page(relative["nodes"], []), staging_root),
           "a staging path and its repository-relative form are one attribution")


def test_duplicate_node_ids_refused() -> None:
    """R10-G01: an id that appears twice must not silently replace itself before comparison.

    Every comparison in the derivation gate keys nodes by id. Two extraction nodes sharing an id
    meant the second replaced the first before anything was compared, so a graph carrying only
    the second passed as derived from the whole extraction. Probed against the live checker,
    that was accepted.

    Measured over the real artifacts first: the extraction repeats 6 ids and every repeat is a
    byte-identical copy, while the graph and the page repeat none. So identical copies stay
    allowed in the extraction, and anything else is refused.
    """
    print("R10-G01 duplicate node ids are refused")
    builder = load("build_public_map_duplicates", "build-public-map.py")
    bind = builder.ensure_artifacts_share_one_run

    def page(nodes, edges=None):
        return ("const RAW_NODES = " + json.dumps(nodes) + ";" + chr(10)
                + "const RAW_EDGES = " + json.dumps(edges or []) + ";")

    node = {"id": "n1", "label": "B.cs", "source_file": "Core/B.cs"}
    graph_node = dict(node, community="0")
    drawn = [{"id": "n1", "label": "B.cs"}]

    allows(lambda: bind({"nodes": [node], "edges": []},
                        {"nodes": [graph_node], "links": []}, page(drawn)),
           "one node per id is accepted")

    allows(lambda: bind({"nodes": [node, dict(node)], "edges": []},
                        {"nodes": [graph_node], "links": []}, page(drawn)),
           "an extraction repeating a node identically is accepted")

    other = {"id": "n1", "label": "C.cs", "source_file": "Core/C.cs"}
    refuses(lambda: bind({"nodes": [node, other], "edges": []},
                         {"nodes": [dict(other, community="0")], "links": []},
                         page([{"id": "n1", "label": "C.cs"}])),
            "an extraction holding one id with two different nodes is refused")

    refuses(lambda: bind({"nodes": [node], "edges": []},
                         {"nodes": [graph_node, dict(graph_node)], "links": []}, page(drawn)),
            "a graph repeating a node id is refused")

    refuses(lambda: bind({"nodes": [node], "edges": []},
                         {"nodes": [graph_node], "links": []}, page(drawn + drawn)),
            "a page repeating a node id is refused")

    refuses(lambda: bind({"nodes": [{"id": "", "label": "B.cs"}], "edges": []},
                         {"nodes": [{"id": "", "label": "B.cs", "community": "0"}], "links": []},
                         page([{"id": "", "label": "B.cs"}])),
            "a node with no id at all is refused")


def test_dropped_duplicate_edge_is_caught_by_provenance() -> None:
    """R8-G04: a duplicate extraction edge removed after the build changes the recorded hash.

    The derivation gate cannot see it — the graph draws one edge per pair whether the extraction
    offered it once or twice, which is what "structurally unobservable" meant. The provenance
    chain is a different mechanism: it records the extraction's bytes, so removing one of two
    identical-key edges changes the digest and the gate refuses.

    Measured on the real artifacts before this was written: dropping one copy of a duplicated
    edge left the derivation gate accepting and changed `extraction_sha256`.
    """
    print("R8-G04 a dropped duplicate edge changes the recorded extraction digest")
    builder = load("build_public_map_dupedge", "build-public-map.py")

    extraction = {
        "nodes": [{"id": "a", "label": "A.cs"}, {"id": "b", "label": "B.cs"}],
        "edges": [
            {"source": "a", "target": "b", "relation": "references", "confidence": "EXTRACTED",
             "source_location": "A.cs:10"},
            {"source": "a", "target": "b", "relation": "references", "confidence": "EXTRACTED",
             "source_location": "A.cs:20"},
        ],
    }

    # Two edges with the same key and different origins: the shape the real extraction carries
    # 462 times. The graph keeps one of them, legitimately.
    graph = {
        "nodes": [{"id": "a", "label": "A.cs", "community": "0"},
                  {"id": "b", "label": "B.cs", "community": "1"}],
        "links": [{"source": "a", "target": "b", "relation": "references",
                   "confidence": "EXTRACTED", "weight": 1.0}],
    }
    page = ("const RAW_NODES = " + json.dumps([{"id": "a", "label": "A.cs"},
                                               {"id": "b", "label": "B.cs"}]) + ";" + chr(10)
            + "const RAW_EDGES = " + json.dumps([{"from": "a", "to": "b", "label": "references",
                                                  "confidence": "EXTRACTED"}]) + ";")

    allows(lambda: builder.ensure_artifacts_share_one_run(extraction, graph, page),
           "an extraction offering one relation twice is accepted, and the graph draws it once")

    dropped = {"nodes": extraction["nodes"], "edges": extraction["edges"][:1]}
    allows(lambda: builder.ensure_artifacts_share_one_run(dropped, graph, page),
           "the derivation gate cannot see the dropped copy - it compares presence, not count")

    # What can see it: the bytes.
    full = hashlib.sha256(json.dumps(extraction, sort_keys=True).encode("utf-8")).hexdigest()
    short = hashlib.sha256(json.dumps(dropped, sort_keys=True).encode("utf-8")).hexdigest()
    check("dropping a duplicate edge changes the extraction digest the metadata records",
          full != short)


def test_porcelain_records() -> None:
    """R4-G03: a rename's source path is consumed whichever status column carries the R."""
    print("R4-G03 porcelain records are parsed by both status columns")
    builder = load("build_public_map_porcelain", "build-public-map.py")
    nul = chr(0)

    def records(*fields):
        return nul.join(fields) + nul

    unstaged_rename = records(" R Helpers/New.cs", "Helpers/Old.cs", " M README.md")
    paths = builder.porcelain_paths(unstaged_rename)
    check("an unstaged rename reports its source path", "Helpers/Old.cs" in paths)
    check("the source path is not read as a status record", "pers/Old.cs" not in paths)
    check("the record after the rename is still read", "README.md" in paths)

    corpus = frozenset()
    cases = (
        ("an unstaged corpus rename out of the corpus",
         records(" R Tests/Moved.cs", "Helpers/Word/Mover.cs"), True),
        ("a staged corpus rename",
         records("R  Helpers/Word/New.cs", "Helpers/Word/Old.cs"), True),
        ("a corpus copy", records("C  Helpers/Word/Copy.cs", "Helpers/Word/Source.cs"), True),
        ("a deleted corpus file", records(" D Helpers/Word/Gone.cs"), True),
        ("an edited corpus document", records(" M README.md"), True),
        ("a rename that touches no corpus path",
         records(" R Tests/New.cs", "Tests/Old.cs"), False),
        ("an untracked scratch file", records("?? notes.txt"), False),
    )
    for description, porcelain, expected in cases:
        verdict = builder.corpus_is_dirty(builder.porcelain_paths(porcelain), corpus)
        check(f"{description} is {'dirty' if expected else 'clean'}", verdict == expected)


def link_outside(link: Path, target: Path) -> bool:
    """Points a repository path at something outside it, however this platform allows.

    <param name="link">Where the link should appear.</param>
    <param name="target">The directory outside the repository it should point at.</param>
    <returns>True when a link was created.</returns>
    """
    try:
        link.symlink_to(target, target_is_directory=True)
        return True
    except (OSError, NotImplementedError):
        pass

    # Windows refuses symlinks without elevation but allows directory junctions, which follow
    # exactly the same way as far as reading files is concerned.
    if os.name != "nt":
        return False

    made = subprocess.run(["cmd", "/c", "mklink", "/J", str(link), str(target)],
                          capture_output=True, text=True, encoding="utf-8", errors="replace")
    return made.returncode == 0 and link.exists()


def test_corpus_never_reaches_outside_the_repository() -> None:
    """R7-G04: the allowlist is applied to a name, but reading follows the link."""
    print("R7-G04 corpus content stays inside the repository")
    allowlist = load("corpus_allowlist_links", "corpus_allowlist.py")

    with tempfile.TemporaryDirectory() as tmp:
        repo = Path(tmp) / "repo"
        (repo / "Core").mkdir(parents=True)
        (repo / "Core" / "Real.cs").write_text("// real", encoding="utf-8")

        outside = Path(tmp) / "outside"
        outside.mkdir()
        (outside / "Secret.cs").write_text("// SECRET", encoding="utf-8")

        if not link_outside(repo / "Core" / "linked", outside):
            check("SKIPPED: this platform will not create a link to test with", True)
            return

        chosen = allowlist.selected_files(repo)
        check("a real file inside the repository is still corpus content",
              "Core/Real.cs" in chosen)
        check("a file reached through a link out of the repository is not corpus content",
              not [c for c in chosen if "linked" in c])
        check("the outside file is not readable as corpus content",
              not allowlist.is_inside_repository(repo / "Core" / "linked" / "Secret.cs", repo))

        # R8-G02: staying inside the repository was not enough. `Core/into-tests -> Tests/` keeps
        # the target in the repository while moving it out of the directory the deny list
        # describes, so `Core/into-tests/Secret.cs` was accepted and the whole of Tests/ went into
        # the public corpus.
        (repo / "Tests").mkdir()
        (repo / "Tests" / "Secret.cs").write_text("// SECRET", encoding="utf-8")
        if link_outside(repo / "Core" / "into-tests", repo / "Tests"):
            chosen = allowlist.selected_files(repo)
            check("a link into a denied directory is not corpus content",
                  not [c for c in chosen if "into-tests" in c])
            check("the denied file behind it is not readable either",
                  not allowlist.is_inside_repository(
                      repo / "Core" / "into-tests" / "Secret.cs", repo))

        # A link whose target is gone must be refused rather than raise, and a link swapped for a
        # different target between selection and copy must not be followed either.
        (outside / "Secret.cs").unlink()
        outside.rmdir()
        check("a dangling link is refused rather than crashing the scan",
              not allowlist.is_inside_repository(repo / "Core" / "linked" / "Secret.cs", repo))
        check("the scan still completes with a dangling link present",
              allowlist.selected_files(repo) == ["Core/Real.cs"])


def test_deletion_is_never_self_authorised() -> None:
    """R7-G02: the proof a directory offers for its own deletion lives inside it."""
    print("R7-G02 deletion is never authorised from inside the target")
    stage = load("stage_corpus_ownership", "stage-corpus.py")
    allowlist = load("corpus_allowlist_ownership", "corpus_allowlist.py")
    run = "0123456789abcdef0123456789abcdef"

    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)

        # A file the manifest names, whose bytes were changed afterwards. Matching on the path
        # alone counted it as accounted-for, so --force would have deleted it (R7-G02).
        edited = stage_a_directory(stage, root, "edited")
        manifest = json.loads((edited / "corpus-manifest.json").read_text(encoding="utf-8"))
        (edited / "Program.cs").write_text("// somebody else's work", encoding="utf-8")
        check("a manifest file whose bytes changed is no longer accounted for",
              allowlist.unlisted_files(edited, manifest["files"], run_id=run) == ["Program.cs"])

        # A sidecar name in a subdirectory is a user's own file, not the root sidecar.
        nested = stage_a_directory(stage, root, "nested")
        (nested / "Core").mkdir()
        (nested / "Core" / "corpus-manifest.json").write_text("{}", encoding="utf-8")
        nested_manifest = json.loads((nested / "corpus-manifest.json").read_text(encoding="utf-8"))
        check("a sidecar name in a subdirectory is not exempt",
              "Core/corpus-manifest.json"
              in allowlist.unlisted_files(nested, nested_manifest["files"], run_id=run))

        # And the whole class: a consistent, well-formed staging directory still cannot authorise
        # its own recursive deletion, because every part of that proof is inside it.
        consistent = stage_a_directory(stage, root, "consistent")
        refuses(lambda: stage.ensure_safe_staging_target(consistent),
                "a fully consistent staging directory still cannot authorise its own deletion")
        allows(lambda: stage.ensure_safe_staging_target(root / "does-not-exist"),
               "a directory that does not exist yet is still allowed")


def test_corpus_cannot_choose_the_interpreter() -> None:
    """R7-G01: nothing the corpus supplies is executed, and nothing is read before it is verified."""
    print("R7-G01 the corpus does not get to choose the interpreter")
    import os

    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        work = root / "hostile" / "graphify-out"
        work.mkdir(parents=True)
        canary = root / "canary.txt"

        # A "python" that does nothing but prove it ran. The publisher used to hand this to
        # subprocess as argv[0] before it had verified anything about the corpus at all.
        if os.name == "nt":
            fake = work / "fake-python.cmd"
            fake.write_text(f'@echo off{chr(13)}{chr(10)}> "{canary}" echo executed{chr(13)}{chr(10)}',
                            encoding="utf-8")
        else:
            fake = work / "fake-python.sh"
            fake.write_text(f"#!/bin/sh{chr(10)}echo executed > '{canary}'{chr(10)}", encoding="utf-8")
            fake.chmod(0o755)

        (work / ".graphify_python").write_text(str(fake), encoding="utf-8")
        (work / ".graphify_extract.json").write_text('{"nodes": [], "edges": []}', encoding="utf-8")

        result = subprocess.run(
            [sys.executable, str(HERE / "build-public-map.py"), "--corpus", str(root / "hostile"),
             "--out", "docs/architecture-map", "--spec", str(HERE / "self-test.py")],
            cwd=REPO_ROOT, capture_output=True, text=True, encoding="utf-8",
            errors="replace")

        check("a corpus with no manifest is refused", result.returncode != 0)
        check("nothing the corpus supplied was executed", not canary.exists())



def junction_command(platform_name: str, link: Path, target: Path) -> list[str] | None:
    """The command that makes a directory link on a platform, or None when it takes no command.

    A function of the platform name rather than of the platform, so both answers are checkable on
    either. R17-G01 was the cost of not doing that: the junction case called `cmd /c mklink /J`
    unconditionally, the Pages workflow runs this file on `ubuntu-latest`, and spawning `cmd` there
    raises FileNotFoundError before any exit code exists to ignore — so the deploy job could not
    get past the fixture at all, and nothing on a Windows machine could notice.
    """
    if platform_name != "nt":
        return None

    return ["cmd", "/c", "mklink", "/J", str(link), str(target)]


def make_directory_link(link: Path, target: Path) -> None:
    """Creates a directory link at `link` pointing to `target`, however this platform makes one.

    What this case is about — a directory link appearing where staging is about to write — is not
    Windows-specific; only the call that makes one is. Windows uses a junction because it needs no
    elevation, and everywhere else uses an ordinary symbolic link.

    Raises when it cannot make one. A link silently not created would leave the case asserting
    nothing, which is the failure this fixture exists to catch elsewhere.
    """
    command = junction_command(os.name, link, target)
    if command is None:
        link.symlink_to(target, target_is_directory=True)
        return

    made = subprocess.run(command, capture_output=True, text=True, encoding="utf-8",
                          errors="replace", check=False)
    if made.returncode != 0 or not link.exists():
        raise OSError(f"could not create a junction at {link}: {made.stderr.strip()}")


def test_directory_links_are_made_without_asking_for_windows() -> None:
    """R17-G01: the Pages job runs this file on Ubuntu, where `cmd` does not exist."""
    link, target = Path("link"), Path("target")

    check("Windows makes a junction, which needs no elevation",
          junction_command("nt", link, target) == ["cmd", "/c", "mklink", "/J", "link", "target"])
    check("no other platform spawns a Windows shell",
          junction_command("posix", link, target) is None)

    # The workflow that would have met this. Named here so a runner change is caught by the
    # fixture rather than by a failed deploy.
    workflow = (REPO_ROOT / ".github" / "workflows" / "deploy-pages.yml").read_text(encoding="utf-8")
    check("the Pages job still runs this self-test before uploading",
          "python graphify/self-test.py" in workflow)
    check("and it still runs on a platform that has no cmd.exe",
          "ubuntu-latest" in workflow)


def test_publish_refuses_a_leaf_that_is_a_link() -> None:
    """R17-G02: the directory's identity was bound; the two files inside it were not.

    `Path.write_text` opens an existing name, so a link left at `index.html` was followed and the
    page written through to whatever it pointed at — and the post-check used `is_file()`, which
    follows a link as well, so it confirmed the write had landed on the link's target.

    Posed with a directory link, because a file symbolic link needs elevation on Windows and a
    fixture that can only skip is how R17-G01 got in. The guard is the same either way: the name
    exists and is not a regular file.
    """
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp).resolve()
        out_dir = root / "architecture-map"
        out_dir.mkdir()
        outside = root / "outside"
        outside.mkdir()
        sentinel = outside / "sentinel.txt"
        sentinel.write_text("not yours", encoding="utf-8")

        make_directory_link(out_dir / "index.html", outside)

        builder = load("build_public_map_leaf", "build-public-map.py")
        refused = False
        try:
            builder.write_regular_file(out_dir / "index.html", "<html>the real page</html>")
        except SystemExit:
            refused = True

        check("a publish leaf that is a link is refused rather than followed", refused)
        check("and what it pointed at is untouched",
              sentinel.read_text(encoding="utf-8") == "not yours")
        check("and nothing was written through it",
              not (outside / "index.html").exists())


def test_publish_writes_a_regular_file() -> None:
    """The control: an ordinary rebuild still replaces the two files in place."""
    with tempfile.TemporaryDirectory() as tmp:
        out_dir = Path(tmp).resolve()
        target = out_dir / "index.html"
        target.write_text("the previous page", encoding="utf-8")

        builder = load("build_public_map_regular", "build-public-map.py")
        builder.write_regular_file(target, "the new page\n")

        check("an existing regular file is replaced",
              target.read_bytes() == b"the new page\n")
        check("no staging file is left behind",
              not (out_dir / "index.html.publishing").exists())


def test_staging_never_deletes() -> None:
    """R8-G03: nothing between the guard and the copy may remove a directory nobody approved."""
    print("R8-G03 staging creates, and never deletes")
    stage = load("stage_corpus_race", "stage-corpus.py")

    # The guard used to be followed by `if out_dir.exists() and any(...): shutil.rmtree(out_dir)`,
    # which re-read neither the guard's decision nor --force. Anything appearing in that window
    # was deleted, so the guard's careful refusals decided nothing.
    # A call, not the word: the comment recording why the delete was removed names it too, and a
    # guard that cannot tell those apart would have to be loosened the first time someone
    # explains themselves.
    source = (HERE / "stage-corpus.py").read_text(encoding="utf-8")
    check("the staging script contains no recursive delete",
          "rmtree(" not in source)

    def stage_into(target: Path) -> bool:
        """Runs the staging script for one target and reports whether it refused."""
        argv = sys.argv
        sys.argv = ["stage-corpus.py", "--out", str(target)]
        try:
            stage.main()
            return False
        except SystemExit as exit_code:
            return bool(exit_code.code)
        finally:
            sys.argv = argv

    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp).resolve()

        # Each case plants the target *after* the guard has already approved it, which is the
        # window the delete used to live in.
        def plant(kind: str):
            def guard(out_dir: Path) -> None:
                out_dir.parent.mkdir(parents=True, exist_ok=True)
                if kind == "file":
                    out_dir.mkdir()
                    (out_dir / "someone-elses.txt").write_text("keep me", encoding="utf-8")
                elif kind == "directory":
                    (out_dir / "sub").mkdir(parents=True)
                    (out_dir / "sub" / "notes.md").write_text("mine", encoding="utf-8")
                else:
                    out_dir.mkdir()
                    make_directory_link(out_dir / "linked", root / "outside")
            return guard

        (root / "outside").mkdir()
        (root / "outside" / "secret.txt").write_text("not yours", encoding="utf-8")

        original = stage.ensure_safe_staging_target
        try:
            for kind in ("file", "directory", "link"):
                target = root / f"race-{kind}"
                stage.ensure_safe_staging_target = plant(kind)
                refused = stage_into(target)

                survived = target.exists() and any(target.iterdir())
                outside_intact = (root / "outside" / "secret.txt").is_file()
                check(f"a {kind} appearing after the guard is refused, not deleted",
                      refused and survived and outside_intact)
        finally:
            stage.ensure_safe_staging_target = original

        # And the ordinary path still works: a name that does not exist yet gets created.
        fresh = root / "fresh"
        check("staging a new directory still succeeds", stage_into(fresh) is False and fresh.is_dir())



def test_publish_target_is_bounded() -> None:
    """R8-G05: the page and the metadata can only ever be written inside docs/."""
    print("R8-G05 publish target is bounded")
    builder = load("build_public_map_out", "build-public-map.py")
    resolve = builder.resolve_publish_directory

    allows(lambda: resolve("docs/architecture-map"), "the documented target is accepted")
    allows(lambda: resolve("docs/architecture-map/nested"), "a subdirectory of it is accepted")

    with tempfile.TemporaryDirectory() as tmp:
        outside = Path(tmp).resolve() / "published-elsewhere"

        # An absolute path was joined to the repository root and resolved, which simply discards
        # the root: the write landed wherever the caller pointed.
        refuses(lambda: resolve(str(outside)), "an absolute path outside the repository is refused")
        check("no directory was created for the refused absolute path", not outside.exists())

        refuses(lambda: resolve("../outside-the-repo"), "a parent-relative path is refused")
        refuses(lambda: resolve("Tests"), "a directory outside docs/ is refused")
        refuses(lambda: resolve("docs/../Tests"), "a path that climbs out of docs/ is refused")

        # A link on the way to the target sends the write somewhere else, whatever the name says.
        linked = REPO_ROOT / "docs" / "architecture-map-link-probe"
        if link_outside(linked, Path(tmp)):
            try:
                refuses(lambda: resolve("docs/architecture-map-link-probe"),
                        "a link standing in for the publish directory is refused")
            finally:
                linked.rmdir()

    # And nothing is created just by asking: the directory appears only once the artifacts have
    # been checked, so a refused build leaves no empty directory behind.
    probe = REPO_ROOT / "docs" / "architecture-map-unused-probe"
    resolve("docs/architecture-map-unused-probe")
    check("resolving a target does not create it", not probe.exists())


def test_staging_rechecks_what_it_actually_reads() -> None:
    """R9-G02: the containment check answers for a moment, the read happens in the next one."""
    print("R9-G02 staging reads what it checked")
    stager = load("stage_corpus_toctou", "stage-corpus.py")

    with tempfile.TemporaryDirectory() as tmp:
        repo = Path(tmp) / "repo"
        (repo / "Core").mkdir(parents=True)
        target = repo / "Core" / "Swapped.cs"
        target.write_text("// approved content", encoding="utf-8")

        outside = Path(tmp) / "outside"
        outside.mkdir()
        secret = outside / "Secret.cs"
        secret.write_text("// SECRET", encoding="utf-8")

        staging = Path(tmp) / "staged"
        staging.mkdir()
        destination = staging / "Swapped.cs"

        # The honest case first, or the refusal below proves nothing about the swap.
        stager.copy_verified(target, destination, "Core/Swapped.cs")
        check("an unswapped file is staged with its own bytes",
              destination.read_text(encoding="utf-8") == "// approved content")
        destination.unlink()

        # The seam is what makes the window reachable at all: it fires after the containment check
        # has returned and before anything is read, which is exactly where another local process
        # would act.
        swapped_to_directory = repo / "Core" / "AsDirectory.cs"
        swapped_to_directory.write_text("// approved content", encoding="utf-8")

        # The window, acted on: by the time copy_verified runs, the name denotes something else.
        swapped_to_directory.unlink()
        swapped_to_directory.mkdir()

        refused_directory = False
        try:
            stager.copy_verified(swapped_to_directory, staging / "AsDirectory.cs",
                                 "Core/AsDirectory.cs")
        except (SystemExit, OSError, PermissionError):
            refused_directory = True

        check("a name that became a directory in the window is not staged as a file",
              refused_directory and not (staging / "AsDirectory.cs").is_file())

        removed = repo / "Core" / "Removed.cs"
        removed.write_text("// approved content", encoding="utf-8")

        removed.unlink()

        refused_missing = False
        try:
            stager.copy_verified(removed, staging / "Removed.cs", "Core/Removed.cs")
        except (SystemExit, OSError):
            refused_missing = True

        check("a name that disappeared in the window is refused rather than half-staged",
              refused_missing and not (staging / "Removed.cs").exists())

        # The seam production leaves alone, so a fixture can drive the window in a full run.
        check("the staging copy exposes a seam for that window, unset in production",
              hasattr(stager, "BEFORE_COPY") and stager.BEFORE_COPY is None)

        # And the case the window exists for, where the platform allows it to be built.
        target.unlink()
        if not link_outside(target, secret):
            check("SKIPPED: this platform will not create a link to test the swap with", True)
            return

        refused = False
        try:
            stager.copy_verified(target, destination, "Core/Swapped.cs")
        except SystemExit:
            refused = True

        check("a name swapped for a link between the check and the read is refused", refused)
        check("no bytes from outside the repository reached the staging directory",
              not destination.exists()
              or "SECRET" not in destination.read_text(encoding="utf-8", errors="replace"))


def test_public_extraction_has_no_dangling_edges() -> None:
    """R30-G01: expected external imports are explicit omissions, not corrupt graph edges."""
    print("R30-G01 public extraction dangling-edge policy")
    try:
        prepare = load("prepare_public_extraction", "prepare-public-extraction.py")
    except FileNotFoundError:
        check("the public extraction normalizer exists", False)
        return

    nodes = [
        {"id": "source", "label": "Source.cs", "file_type": "code",
         "source_file": "Core/Source.cs", "_origin": "ast"},
        {"id": "target", "label": "Target", "file_type": "code",
         "source_file": "Core/Target.cs", "_origin": "ast"},
    ]
    external_import = {
        "source": "source", "target": "system_text", "relation": "imports",
        "confidence": "EXTRACTED", "source_file": "Core/Source.cs", "_origin": "ast",
        "metadata": {"using_kind": "namespace", "target_fqn": "System.Text"},
    }
    real_edge = {
        "source": "source", "target": "target", "relation": "calls",
        "confidence": "EXTRACTED", "source_file": "Core/Source.cs", "_origin": "ast",
    }

    cleaned, removed = prepare.normalise({"nodes": nodes, "edges": [external_import, real_edge]})
    check("an AST import to an external namespace is removed before graph construction",
          removed == 1 and cleaned["edges"] == [real_edge])

    unresolved_call = dict(real_edge, target="missing")
    refuses(lambda: prepare.normalise({"nodes": nodes, "edges": [unresolved_call]}),
            "a non-import dangling edge fails closed instead of disappearing")

    semantic_import = dict(external_import, _origin="semantic")
    refuses(lambda: prepare.normalise({"nodes": nodes, "edges": [semantic_import]}),
            "a semantic dangling import is not mistaken for an expected AST external dependency")

    ast_file = {
        "id": "core_licensemanager", "label": "LicenseManager.cs", "file_type": "code",
        "source_file": "Core/LicenseManager.cs", "source_location": "L1", "_origin": "ast",
    }
    mentioned_file = {
        "id": "docs_developers_license_manager", "label": "LicenseManager.cs",
        "file_type": "concept", "source_file": "docs/developers.html",
    }
    mention_edge = {
        "source": "docs_developers_page", "target": mentioned_file["id"],
        "relation": "references", "confidence": "EXTRACTED",
        "source_file": "docs/developers.html",
    }
    mention_nodes = [
        ast_file,
        mentioned_file,
        {"id": "docs_developers_page", "label": "Developer guide",
         "file_type": "document", "source_file": "docs/developers.html"},
    ]
    cleaned, removed = prepare.normalise({
        "nodes": mention_nodes,
        "edges": [mention_edge],
        "hyperedges": [{"id": "guide", "nodes": [mentioned_file["id"]]}],
    })
    check("a document mention of one uniquely named AST file is normalized before build",
          removed == 0
          and mentioned_file["id"] not in {node["id"] for node in cleaned["nodes"]}
          and cleaned["edges"][0]["target"] == ast_file["id"]
          and cleaned["hyperedges"][0]["nodes"] == [ast_file["id"]]
          and cleaned["_graphify_public_normalization"]
                     ["normalized_semantic_file_mentions"] == 1)

    second_ast_file = dict(ast_file, id="legacy_licensemanager",
                           source_file="Legacy/LicenseManager.cs")
    ambiguous, _ = prepare.normalise({
        "nodes": mention_nodes + [second_ast_file], "edges": [mention_edge],
    })
    check("an ambiguous same-named AST file does not absorb a document concept",
          mentioned_file["id"] in {node["id"] for node in ambiguous["nodes"]}
          and ambiguous["edges"][0]["target"] == mentioned_file["id"])

    builder = load("build_public_map_health", "build-public-map.py")
    if not hasattr(builder, "ensure_extraction_health"):
        check("the public builder requires extraction health evidence", False)
        return
    healthy = {
        "node_count": 2,
        "raw_edge_count": 1,
        "dangling_endpoint_edges": 0,
        "missing_endpoint_edges": 0,
        "self_loop_edges": 0,
        "omitted_external_import_edges": 1,
        "normalized_semantic_file_mentions": 0,
    }
    allows(lambda: builder.ensure_extraction_health(healthy),
           "zero-dangling diagnostics are publishable")
    refuses(lambda: builder.ensure_extraction_health(
                dict(healthy, dangling_endpoint_edges=1)),
            "a public extraction with a dangling endpoint is refused")
    refuses(lambda: builder.ensure_extraction_health({}),
            "missing extraction diagnostics fail closed")



def main() -> int:
    test_staging_guard()
    test_directory_links_are_made_without_asking_for_windows()
    test_publish_refuses_a_leaf_that_is_a_link()
    test_publish_writes_a_regular_file()
    test_staging_never_deletes()
    test_publish_target_is_bounded()
    test_staged_manifest_binds_bytes()
    test_artifacts_share_one_run()
    test_page_hardening()
    test_leak_scanner()
    test_leak_scanner_scopes()
    test_workflow_triggers()
    test_artifacts_derive_from_each_other()
    test_lossless_relationship_artifact()
    test_deleted_sources_count_as_dirty()
    test_work_directory_is_accounted_for()
    test_page_payload_matches()
    test_artifact_forgeries_refused()
    test_duplicate_node_ids_refused()
    test_dropped_duplicate_edge_is_caught_by_provenance()
    test_porcelain_records()
    test_corpus_never_reaches_outside_the_repository()
    test_staging_rechecks_what_it_actually_reads()
    test_public_extraction_has_no_dangling_edges()
    test_deletion_is_never_self_authorised()
    test_corpus_cannot_choose_the_interpreter()
    test_gate_fails_closed()

    print()
    if failures:
        print(f"{len(failures)} of {checks} fixtures failed:")
        for name in failures:
            print(f"  - {name}")
        return 1
    print(f"All {checks} fixtures passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
