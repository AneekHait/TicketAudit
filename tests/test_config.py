"""
Tests for config.py — ConfigManager.
"""
import json
import os
import pathlib
import pytest
import tempfile
from config import ConfigManager, DEFAULT_CONFIG


class TestConfigManagerLoad:
    def test_loads_valid_json(self, tmp_path):
        cfg_file = tmp_path / "config.json"
        cfg_file.write_text(json.dumps({"null_threshold": 30.0}))
        mgr = ConfigManager(str(cfg_file))
        assert mgr.get("null_threshold") == 30.0

    def test_falls_back_to_defaults_on_missing_file(self, tmp_path):
        missing = str(tmp_path / "nonexistent.json")
        mgr = ConfigManager(missing)
        assert mgr.get("null_threshold") == DEFAULT_CONFIG["null_threshold"]

    def test_falls_back_to_defaults_on_malformed_json(self, tmp_path):
        cfg_file = tmp_path / "config.json"
        cfg_file.write_text("{this is not valid json}")
        mgr = ConfigManager(str(cfg_file))
        assert mgr.get("null_threshold") == DEFAULT_CONFIG["null_threshold"]

    def test_missing_keys_filled_with_defaults(self, tmp_path):
        cfg_file = tmp_path / "config.json"
        cfg_file.write_text(json.dumps({"null_threshold": 25.0}))
        mgr = ConfigManager(str(cfg_file))
        assert mgr.get("null_threshold") == 25.0
        assert mgr.get("sidebar_width") == DEFAULT_CONFIG["sidebar_width"]


class TestConfigManagerSave:
    def test_round_trip(self, tmp_path):
        cfg_file = str(tmp_path / "config.json")
        mgr = ConfigManager(cfg_file)
        mgr.set("null_threshold", 42.0)
        mgr.save()

        mgr2 = ConfigManager(cfg_file)
        assert mgr2.get("null_threshold") == 42.0

    def test_update_multiple_keys(self, tmp_path):
        cfg_file = str(tmp_path / "config.json")
        mgr = ConfigManager(cfg_file)
        mgr.update({"null_threshold": 15.0, "sidebar_width": 300})
        assert mgr.get("null_threshold") == 15.0
        assert mgr.get("sidebar_width") == 300


class TestConfigFilePath:
    def test_default_config_path_is_absolute(self):
        """CONFIG_FILE must be absolute so it resolves correctly regardless of CWD."""
        from config import CONFIG_FILE
        assert os.path.isabs(CONFIG_FILE), (
            f"CONFIG_FILE should be absolute, got: {CONFIG_FILE}"
        )

    def test_no_use_polars_in_default_config(self):
        """use_polars was removed as a dead key — it must not appear in DEFAULT_CONFIG."""
        assert "use_polars" not in DEFAULT_CONFIG
