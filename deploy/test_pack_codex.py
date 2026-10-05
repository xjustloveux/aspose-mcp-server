"""Exercise the custom-marketplace archive without external dependencies."""

import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile


SCRIPT = Path(__file__).with_name("pack-codex.py")


class CodexPackagingTests(unittest.TestCase):
    """Check installable archives, input rejection and preservation of user files."""

    def setUp(self):
        """Create an isolated workspace fixture; propagate filesystem failures."""
        # Keep fixtures under the ignored workspace temp/ directory for inspection.
        temp_root = Path(__file__).resolve().parents[1] / "temp"
        temp_root.mkdir(exist_ok=True)
        self.root = Path(tempfile.mkdtemp(prefix="codex-pack-", dir=temp_root))
        self.plugin = self.root / "plugins" / "aspose-mcp-server"
        self.plugin.mkdir(parents=True)
        catalog = self.root / ".agents" / "plugins"
        catalog.mkdir(parents=True)
        (catalog / "marketplace.json").write_text(json.dumps({
            "name": "aspose-local",
            "plugins": [{"name": "aspose-mcp-server", "source": {
                "source": "local", "path": "./plugins/aspose-mcp-server"}}],
        }), encoding="utf-8")
        (self.plugin / "plugin.json").write_text(json.dumps({
            "name": "aspose-mcp-server", "version": "0.1.0",
        }), encoding="utf-8")
        compatibility = self.plugin / ".codex-plugin"
        compatibility.mkdir()
        (compatibility / "plugin.json").write_text(json.dumps({
            "name": "aspose-mcp-server", "version": "0.1.0",
            "mcpServers": "./mcp.json",
        }), encoding="utf-8")
        (self.plugin / "mcp.json").write_text(json.dumps({"mcpServers": {
            "aspose": {"type": "stdio", "command": "AsposeMcpServer",
                       "args": ["--stdio"]}}}), encoding="utf-8")
        (self.plugin / "README.md").write_text("Installation instructions", encoding="utf-8")
        (self.root / "LICENSE").write_text("MIT", encoding="utf-8")
        (self.plugin / "Aspose.Total.lic").write_text("private license", encoding="utf-8")
        (self.plugin / "secrets.json").write_text("private configuration", encoding="utf-8")
        self.output = self.root / "marketplace.zip"

    def pack(self, version="1.2.3", archive_format="marketplace"):
        """Run the packer with the supplied release version input.

        Args:
            version: Version input, defaulting to 1.2.3. Negative tests may supply
                an invalid string and expect the packer to reject it.
            archive_format: marketplace (default) or plugin archive layout.

        Returns:
            subprocess.CompletedProcess containing the exit code and captured output.

        Raises:
            AssertionError: The packer has not been implemented.
            OSError: The Python subprocess cannot be started.
        """
        self.assertTrue(SCRIPT.is_file(), "Codex marketplace packer has not been implemented")
        return subprocess.run([
            sys.executable, str(SCRIPT), "--repo-root", str(self.root),
            "--version", version, "--output", str(self.output),
            "--format", archive_format,
        ], capture_output=True, text=True, check=False)

    def test_archive_is_installable_and_contains_only_distribution_files(self):
        """Resolve the catalog inside the ZIP and check both manifest versions."""
        result = self.pack()
        self.assertEqual(0, result.returncode, result.stderr)
        with zipfile.ZipFile(self.output) as archive:
            self.assertEqual({
                ".agents/plugins/marketplace.json",
                "plugins/aspose-mcp-server/plugin.json",
                "plugins/aspose-mcp-server/.codex-plugin/plugin.json",
                "plugins/aspose-mcp-server/mcp.json",
                "plugins/aspose-mcp-server/README.md",
                "plugins/aspose-mcp-server/LICENSE",
            }, set(archive.namelist()))
            manifest = json.loads(archive.read("plugins/aspose-mcp-server/plugin.json"))
            self.assertEqual("1.2.3", manifest["version"])
            legacy = json.loads(archive.read("plugins/aspose-mcp-server/.codex-plugin/plugin.json"))
            self.assertEqual(manifest["version"], legacy["version"])
            self.assertEqual(manifest["name"], legacy["name"])
            self.assertIn("plugins/aspose-mcp-server/" + legacy["mcpServers"][2:], archive.namelist())
            catalog = json.loads(archive.read(".agents/plugins/marketplace.json"))
            plugin_root = catalog["plugins"][0]["source"]["path"][2:]
            self.assertIn(plugin_root + "/plugin.json", archive.namelist())
        source = json.loads((self.plugin / "plugin.json").read_text(encoding="utf-8"))
        self.assertEqual("0.1.0", source["version"])

    def test_plugin_archive_has_root_manifests_and_needs_no_catalog(self):
        """Check a single-plugin ZIP excludes the catalog and private files."""
        (self.root / ".agents/plugins/marketplace.json").write_text("invalid catalog", encoding="utf-8")
        result = self.pack(archive_format="plugin")
        self.assertEqual(0, result.returncode, result.stderr)
        with zipfile.ZipFile(self.output) as archive:
            self.assertEqual({"plugin.json", ".codex-plugin/plugin.json", "mcp.json",
                              "README.md", "LICENSE"}, set(archive.namelist()))
            manifest = json.loads(archive.read("plugin.json"))
            legacy = json.loads(archive.read(".codex-plugin/plugin.json"))
            self.assertEqual("1.2.3", manifest["version"])
            self.assertEqual(manifest["version"], legacy["version"])
            self.assertIn(legacy["mcpServers"][2:], archive.namelist())
        source = json.loads((self.plugin / "plugin.json").read_text(encoding="utf-8"))
        self.assertEqual("0.1.0", source["version"])

    def test_plugin_existing_output_is_preserved(self):
        """Keep an existing file intact when packaging a single plugin."""
        self.output.write_bytes(b"existing plugin archive")
        result = self.pack(archive_format="plugin")
        self.assertNotEqual(0, result.returncode)
        self.assertEqual(b"existing plugin archive", self.output.read_bytes())

    def test_codex_only_server_fields_are_rejected_before_packaging(self):
        """Prevent native Codex options from invalidating a portable MCP server."""
        config_path = self.plugin / "mcp.json"
        config = json.loads(config_path.read_text(encoding="utf-8"))
        config["mcpServers"]["aspose"]["env_vars"] = ["ASPOSE_LICENSE_PATH"]
        config_path.write_text(json.dumps(config), encoding="utf-8")
        for archive_format in ("marketplace", "plugin"):
            with self.subTest(archive_format=archive_format):
                result = self.pack(archive_format=archive_format)
                self.assertNotEqual(0, result.returncode)
                self.assertIn("env_vars", result.stderr)
                self.assertFalse(self.output.exists())

    def test_repository_mcp_uses_only_portable_stdio_fields(self):
        """Check the shipped server against the closed Agent Plugins stdio field set."""
        config_path = SCRIPT.parent.parent / "plugins/aspose-mcp-server/mcp.json"
        config = json.loads(config_path.read_text(encoding="utf-8"))
        server = config["mcpServers"]["aspose"]
        self.assertEqual("stdio", server["type"])
        self.assertEqual(set(), set(server) - {"type", "command", "args", "env", "cwd"})

    def test_existing_output_is_preserved(self):
        """Verify failure leaves an existing archive byte-for-byte intact."""
        self.output.write_bytes(b"existing archive")
        result = self.pack()
        self.assertNotEqual(0, result.returncode)
        self.assertEqual(b"existing archive", self.output.read_bytes())

    def test_invalid_version_creates_no_archive(self):
        """Reject paths, leading zeros and Unicode digits in both archive formats."""
        for archive_format in ("marketplace", "plugin"):
            for version in ("../../outside", "01.2.3", "1\u0662.2.3", "1.2\u0663.3", "1.2.3\u0664"):
                with self.subTest(archive_format=archive_format, version=version):
                    result = self.pack(version, archive_format)
                    self.assertNotEqual(0, result.returncode)
                    self.assertFalse(self.output.exists())

    def test_missing_manifest_creates_no_archive(self):
        """Reject incomplete sources before creating an archive."""
        # Use a separate incomplete fixture without removing any file.
        empty_root = self.root / "incomplete"
        empty_root.mkdir()
        original_root = self.root
        self.root = empty_root
        result = self.pack()
        self.root = original_root
        self.assertNotEqual(0, result.returncode)
        self.assertFalse(self.output.exists())


if __name__ == "__main__":
    unittest.main()
