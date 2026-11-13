import contextlib
import json
import zipfile

import pytest
from conftest import make_message

from outlook_mail_archiver.config import Config, ExportConfig, MailArchiverConfig
from outlook_mail_archiver import export as export_mod


class FakeStore:
    DisplayName = "Test Store"


@pytest.fixture
def fake_outlook(monkeypatch):
    registry: dict[str, list] = {}

    @contextlib.contextmanager
    def fake_session(store_name):
        yield None, FakeStore()

    def fake_find(store, spec):
        return spec if spec in registry else None

    def fake_iter(folder, date_from, date_to):
        yield from registry[folder]

    def fake_extract(item):
        return item  # items are already MessageData

    monkeypatch.setattr(export_mod, "outlook_session", fake_session)
    monkeypatch.setattr(export_mod, "find_folder", fake_find)
    monkeypatch.setattr(export_mod, "iter_folder_items", fake_iter)
    monkeypatch.setattr(export_mod, "extract_message", fake_extract)
    return registry


def _config(tmp_path, folders):
    return Config(
        outlook_store="",
        export=ExportConfig(staging_dir=tmp_path, folders=folders),
        mail_archiver=MailArchiverConfig(),
        source_path=tmp_path / "config.toml",
    )


def _manifest(summary):
    return json.loads((summary.package_dir / "manifest.json").read_text("utf-8"))


def _zip_names(summary):
    with zipfile.ZipFile(summary.package_dir / "messages.zip") as zf:
        return zf.namelist()


def test_export_basic(tmp_path, fake_outlook):
    fake_outlook["Inbox"] = [
        make_message(internet_message_id="<a@x>", subject="one"),
        make_message(internet_message_id="<b@x>", subject="two"),
    ]
    summary = export_mod.run_export(_config(tmp_path, ["Inbox"]))
    assert summary.exported == 2 and summary.ok
    assert len(_zip_names(summary)) == 2
    man = _manifest(summary)
    assert man["totals"]["exported"] == 2
    assert man["source_store"] == "Test Store"


def test_slash_folder_maps_to_single_dir(tmp_path, fake_outlook):
    fake_outlook[r"Qlik\/RMM"] = [make_message(internet_message_id="<a@x>")]
    summary = export_mod.run_export(_config(tmp_path, [r"Qlik\/RMM"]))
    names = _zip_names(summary)
    assert all(n.startswith("Qlik_RMM~") for n in names)
    assert all(n.count("/") == 1 for n in names)


def test_unsupported_item_skipped_not_dropped(tmp_path, fake_outlook):
    fake_outlook["Cal"] = [
        make_message(message_class="IPM.Appointment", subject="meeting"),
        make_message(internet_message_id="<a@x>", subject="mail"),
    ]
    summary = export_mod.run_export(_config(tmp_path, ["Cal"]))
    assert summary.exported == 1
    assert summary.skipped_unsupported == 1
    statuses = {m["status"] for m in _manifest(summary)["messages"]}
    assert "skipped-unsupported" in statuses


def test_true_duplicate_item_exported_once(tmp_path, fake_outlook):
    # Same PHYSICAL item (same EntryID) reached via two folders -> exported once.
    fake_outlook["A"] = [make_message(entry_id="SAME", internet_message_id="<a@x>")]
    fake_outlook["B"] = [make_message(entry_id="SAME", internet_message_id="<a@x>")]
    summary = export_mod.run_export(_config(tmp_path, ["A", "B"]))
    assert summary.exported == 1
    assert summary.skipped_duplicate == 1
    assert len(_zip_names(summary)) == 1


def test_message_id_conflict_both_exported(tmp_path, fake_outlook):
    # Regression: two DISTINCT items sharing a Message-ID but with different
    # bodies must BOTH be exported and the conflict flagged — never collapsed.
    fake_outlook["A"] = [make_message(entry_id="E_A", internet_message_id="<dup@x>",
                                      html_body="<p>one</p>")]
    fake_outlook["B"] = [make_message(entry_id="E_B", internet_message_id="<dup@x>",
                                      html_body="<p>two</p>")]
    summary = export_mod.run_export(_config(tmp_path, ["A", "B"]))
    assert summary.exported == 2
    assert len(_zip_names(summary)) == 2
    assert summary.conflicts == 1
    man = _manifest(summary)
    assert man["totals"]["conflicts"] == 1
    assert man["conflicts"][0]["kind"] == "different-body"
    assert set(man["conflicts"][0]["entry_ids"]) == {"E_A", "E_B"}


def test_folder_not_found_is_failure(tmp_path, fake_outlook):
    summary = export_mod.run_export(_config(tmp_path, ["DoesNotExist"]))
    assert summary.failed == 1 and not summary.ok
    assert _manifest(summary)["errors"][0]["reason"] == "folder-not-found"


def test_folder_mapping_is_order_independent(tmp_path, fake_outlook):
    # 'A/B' (one name) and 'A_B' (one name) must map to stable, distinct dirs
    # regardless of selection order.
    fake_outlook[r"A\/B"] = [make_message(internet_message_id="<a@x>")]
    fake_outlook["A_B"] = [make_message(internet_message_id="<b@x>")]

    s1 = export_mod.run_export(_config(tmp_path / "r1", [r"A\/B", "A_B"]))
    s2 = export_mod.run_export(_config(tmp_path / "r2", ["A_B", r"A\/B"]))
    dirs1 = {n.split("/")[0] for n in _zip_names(s1)}
    dirs2 = {n.split("/")[0] for n in _zip_names(s2)}
    assert len(dirs1) == 2, f"folders collided: {dirs1}"
    assert dirs1 == dirs2, "folder mapping changed with selection order"
