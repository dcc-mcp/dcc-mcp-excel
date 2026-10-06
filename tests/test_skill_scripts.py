"""Skill script integration tests — run the scripts like the gateway would.

The gateway executes skills/scripts/<name>.py as a subprocess feeding stdin
JSON (dcc-mcp-core execute_script convention); these tests pin that contract.
The scripts are the same ones a Linux runner executes, so this is the
end-to-end proof that the headless path does not need Excel.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILLS = ROOT / "src/dcc_mcp_excel/skills"
GENERATE = SKILLS / "excel-workbook/scripts/generate_workbook.py"
VALIDATE = SKILLS / "excel-workbook/scripts/validate_workbook.py"
INSPECT = SKILLS / "excel-workbook/scripts/inspect_workbook.py"
CAPABILITIES = SKILLS / "excel-capabilities/scripts/capabilities.py"
PREFLIGHT = SKILLS / "excel-capabilities/scripts/preflight.py"
EXAMPLE = ROOT / "examples/shot_list.json"


def _run_script(script: Path, params: dict, env: dict | None = None) -> tuple[dict, str]:
    proc = subprocess.run(
        [sys.executable, str(script)],
        input=json.dumps(params),
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=300,
        check=False,
        env=env,
    )
    assert proc.returncode == 0, proc.stderr
    return json.loads(proc.stdout), proc.stderr


def test_generate_workbook_script(tmp_path: Path, skill_env: dict) -> None:
    result, _ = _run_script(GENERATE, {"input": str(EXAMPLE), "output_dir": str(tmp_path)}, skill_env)
    assert result["success"], result
    assert result["context"]["backend"] == "openxml"
    assert result["context"]["read_back"]["ok"], result["context"]["read_back"]
    assert Path(result["context"]["artifact"]).is_file()


def test_generate_workbook_script_without_verify(tmp_path: Path, skill_env: dict) -> None:
    result, _ = _run_script(GENERATE, {"input": str(EXAMPLE), "output_dir": str(tmp_path), "verify": False}, skill_env)
    assert result["success"], result
    assert result["context"]["read_back"] is None


def test_generate_workbook_script_reports_failure_on_bad_ir(tmp_path: Path, skill_env: dict) -> None:
    bad = tmp_path / "bad.json"
    bad.write_text(json.dumps({"schema_version": "office-ir/1.0", "kind": "workbook"}), encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, str(GENERATE)],
        input=json.dumps({"input": str(bad), "output_dir": str(tmp_path)}),
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
        env=skill_env,
    )
    assert proc.returncode == 1
    assert json.loads(proc.stdout)["success"] is False


def test_validate_workbook_script(skill_env: dict) -> None:
    result, _ = _run_script(VALIDATE, {"input": str(EXAMPLE)}, skill_env)
    assert result["success"], result
    assert result["context"]["ok"]


def test_validate_artifacts_script(tmp_path: Path, skill_env: dict) -> None:
    result, _ = _run_script(GENERATE, {"input": str(EXAMPLE), "output_dir": str(tmp_path)}, skill_env)
    assert result["success"]
    checked, _ = _run_script(VALIDATE, {"input": str(tmp_path)}, skill_env)
    assert checked["success"], checked


def test_inspect_workbook_script(tmp_path: Path, skill_env: dict) -> None:
    generated, _ = _run_script(GENERATE, {"input": str(EXAMPLE), "output_dir": str(tmp_path)}, skill_env)
    assert generated["success"]
    result, _ = _run_script(INSPECT, {"input": generated["context"]["artifact"]}, skill_env)
    assert result["success"], result
    assert result["context"]["sheet_count"] >= 1
    assert result["context"]["populated_cells"] > 0


def test_capabilities_script(skill_env: dict) -> None:
    result, _ = _run_script(CAPABILITIES, {}, skill_env)
    assert result["success"], result
    assert "workbook.compile" in result["context"]["verified"]
    assert "excel.workbook.calculate" in result["context"]["host_limited"]


def test_preflight_script(skill_env: dict) -> None:
    result, _ = _run_script(PREFLIGHT, {}, skill_env)
    assert isinstance(result["success"], bool)
    assert result["context"]["app"] == "excel"
    assert result["context"]["host_matrix"]
