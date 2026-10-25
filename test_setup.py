"""Setup checks use temporary empty homes and synthetic data only."""
import io
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

import app
from launcher import configured_home, environment_python, prepare_environment, save_home, validate_home


class SetupTests(unittest.TestCase):
    def test_configuration_round_trip_and_environment_priority(self):
        with tempfile.TemporaryDirectory() as folder:
            code = Path(folder) / "code"
            code.mkdir()
            selected = Path(folder) / "Özel veriler"
            save_home(selected, code)
            with patch.dict(os.environ, {"HARCAMA_HOME": ""}):
                self.assertEqual(configured_home(code), str(selected))
            with patch.dict(os.environ, {"HARCAMA_HOME": "explicit-override"}):
                self.assertEqual(configured_home(code), "explicit-override")
            self.assertFalse((code / "home.txt.tmp").exists())
            self.assertFalse(selected.exists(), "Saving a setting must not inspect/create private data")

    def test_rejects_project_folder_and_relative_paths(self):
        with tempfile.TemporaryDirectory() as folder:
            code = Path(folder) / "code"
            for value in ("", "relative", str(code), str(code / "data"), str(code.anchor)):
                with self.subTest(value=value), self.assertRaises(ValueError):
                    validate_home(value, code)
            selected = Path(folder) / "private"
            self.assertEqual(validate_home(str(selected), code), selected.resolve())

    def test_install_failure_is_retryable_and_cache_avoids_download(self):
        with tempfile.TemporaryDirectory() as folder:
            code = Path(folder)
            python = environment_python(code)
            python.parent.mkdir(parents=True)
            python.touch()
            (code / "requirements.txt").write_text("example==1\n")
            failed = subprocess.CompletedProcess([], 1, "", "Connection unavailable")
            passed = subprocess.CompletedProcess([], 0, "", "")
            marker = code / ".venv" / "fintracker-requirements.sha256"
            with patch("launcher.subprocess.run", return_value=failed):
                with self.assertRaisesRegex(RuntimeError, "Connection unavailable"):
                    prepare_environment(lambda _: None, code)
            self.assertFalse(marker.exists())
            with patch("launcher.subprocess.run", return_value=passed) as run:
                self.assertEqual(prepare_environment(lambda _: None, code), python)
                self.assertTrue(any("pip" in call.args[0] for call in run.call_args_list))
            with patch("launcher.subprocess.run", return_value=passed) as run:
                prepare_environment(lambda _: None, code)
                self.assertFalse(any("pip" in call.args[0] for call in run.call_args_list))
            (code / "requirements.txt").write_text("example==2\n")
            with patch("launcher.subprocess.run", return_value=passed) as run:
                prepare_environment(lambda _: None, code)
                self.assertTrue(any("pip" in call.args[0] for call in run.call_args_list))

    def test_empty_home_serves_first_run_without_pipeline(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            handler = app.Handler.__new__(app.Handler)
            handler.path = "/"
            handler._host_ok = lambda: True
            sent = []
            handler._send = lambda *args: sent.append(args)
            with patch.object(app.H, "OUT_HTML", root / "Harcamalarim.html"), \
                 patch.object(app.H, "ST_CSV", root / "data" / "statements.csv"), \
                 patch.object(app.H, "DOCS_DIR", root / "Belgeler"), \
                 patch.object(app, "run_pipeline") as pipeline:
                handler.do_GET()
            pipeline.assert_not_called()
            self.assertEqual(sent[0][0], 200)
            self.assertIn("İlk ekstrelerinizi".encode(), sent[0][1])
            self.assertNotIn(b"/*__APP__*/null", sent[0][1])
            self.assertIn(app.TOKEN.encode(), sent[0][1])

    def test_import_reports_not_ready_and_rejects_missing_token(self):
        with tempfile.TemporaryDirectory() as folder:
            handler = app.Handler.__new__(app.Handler)
            handler.path = "/api/import"
            handler._host_ok = lambda: True
            handler.headers = {"X-Token": app.TOKEN, "Content-Length": "2"}
            handler.rfile = io.BytesIO(b"{}")
            sent = []
            handler._json = lambda *args: sent.append(args)
            with patch.object(app.H, "OUT_HTML", Path(folder) / "absent.html"), \
                 patch.object(app, "run_pipeline", return_value=(True, "Add a card statement")):
                handler.do_POST()
            self.assertEqual(sent[0][0], 200)
            self.assertFalse(sent[0][1]["ready"])
            handler.headers = {}
            handler.do_POST()
            self.assertEqual(sent[-1][0], 403)

    def test_empty_import_does_not_crash_or_create_dashboard(self):
        with tempfile.TemporaryDirectory() as folder:
            result = subprocess.run(
                [sys.executable, str(Path(__file__).with_name("harcama.py")), "--no-recalc"],
                env=dict(os.environ, HARCAMA_HOME=folder, PYTHONIOENCODING="utf-8"),
                capture_output=True, text=True, encoding="utf-8", timeout=30)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn("kredi kartı ekstresi", result.stdout)
            self.assertFalse((Path(folder) / "Harcamalarim.html").exists())


if __name__ == "__main__":
    unittest.main()
