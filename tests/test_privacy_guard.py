import hashlib
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location("check_no_pii", Path(__file__).parents[1] / "scripts" / "check_no_pii.py")
guard = importlib.util.module_from_spec(spec)
spec.loader.exec_module(guard)


def test_guard_flags_real_contact_data():
    # built at runtime so this test file does not trip the guard itself
    row = "Ana," + "ana" + "@" + "gmail.com," + ".".join(["187", "190", "1", "20"])
    issues = guard.problems(Path("x.csv"), "name,email,ip\n" + row)
    assert any("email" in i for i in issues)
    assert any("IPv4" in i for i in issues)


def test_guard_allows_reserved_demo_values():
    assert guard.problems(Path("x.csv"), "demo+0001@example.com,203.0.113.7") == []


def test_guard_flags_blocked_names_by_hash(monkeypatch):
    monkeypatch.setattr(guard, "BLOCKED_NAME_HASHES", {hashlib.sha256(b"acmebrand").hexdigest()})
    assert guard.problems(Path("x.md"), "Results for Acme Brand customers") == ["1 blocked client name(s)"]


def test_guard_flags_key_shaped_strings():
    assert guard.problems(Path("x.py"), 'key = "sk-' + "a" * 30 + '"') == ["API-key shaped string"]
