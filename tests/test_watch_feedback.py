"""Watcher plumbing: diagnosis idempotence and learned-dedupe, offline."""
import watch_feedback as wf


def test_diagnosis_key_distinguishes_constraints():
    assert wf.diagnosis_key({"job_url": "u", "constraint": "a"}) == \
        wf.diagnosis_key({"job_url": "u", "constraint": "a"})
    assert wf.diagnosis_key({"job_url": "u", "constraint": "a"}) != \
        wf.diagnosis_key({"job_url": "u", "constraint": "b"})


def test_load_diagnosed_roundtrip_and_corruption(tmp_path):
    path = str(tmp_path / "diagnosis.jsonl")
    rec = {"job_url": "https://x/j1", "constraint": "not_fresher",
           "root_cause": "jd_unavailable_at_scrape"}
    wf._append_jsonl(path, rec)
    with open(path, "a") as f:
        f.write("garbage\n\n")
    keys = wf.load_diagnosed(path)
    assert keys == {("https://x/j1", "not_fresher")}
    assert wf.load_diagnosed(str(tmp_path / "missing.jsonl")) == set()


def test_learned_entries_dedupe_by_kind_scope_value(tmp_path):
    path = str(tmp_path / "learned.jsonl")
    from saasnews.feedback import learning_key
    entries = [
        {"kind": "location_token", "value": "texas", "scope": "global"},
        {"kind": "location_token", "value": "Texas", "scope": "global"},
        {"kind": "title_negative", "value": r"\bsde[\s\-]*2\b", "scope": "global"},
    ]
    seen = set()
    appended = 0
    for e in entries:
        k = learning_key(e)
        if k in seen:
            continue
        seen.add(k)
        wf._append_jsonl(path, e)
        appended += 1
    assert appended == 2
    from saasnews.feedback import load_learned
    assert len(load_learned(path)) == 2
