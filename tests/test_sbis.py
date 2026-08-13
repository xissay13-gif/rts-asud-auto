import tempfile
import unittest
import json
from pathlib import Path

from flows.sbis import (
    content_from_file,
    derive_surname,
    load_manifest_ids,
    _output_path,
    parse_protocol_filename,
    prepare_items,
    scan_files,
)


class SbisFlowTests(unittest.TestCase):
    def test_output_registry_is_created_in_source_folder(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "Корр Вх"
            root.mkdir()

            output = Path(_output_path(root))

            self.assertEqual(output.parent, root.resolve())
            self.assertEqual(output.name, "СБИС_Корр_Вх_резолюции.xlsx")

    def test_content_uses_filename_without_extension(self):
        path = Path("Ответ на требование.pdf")
        self.assertEqual(content_from_file(path), "Ответ на требование")
        self.assertEqual(content_from_file(path, include_extension=True),
                         "Ответ на требование.pdf")

    def test_surname_from_new_filename_layout(self):
        root = Path("C:/SBIS")
        path = root / "20260731" / "Иванов Иван Иванович Письмо № 2.pdf"
        self.assertEqual(derive_surname(path, root), "Иванов")

    def test_surname_from_old_dated_parent_layout(self):
        root = Path("C:/SBIS")
        path = (root / "КоррВх" / "07.2026"
                / "31.07.2026 Петров Петр Петрович Письмо № 2"
                / "ответ.docx")
        self.assertEqual(derive_surname(path, root), "Петров")

    def test_override_wins(self):
        root = Path("C:/SBIS")
        self.assertEqual(
            derive_surname(root / "непонятное имя.pdf", root, "Сидоров"),
            "Сидоров")

    def test_protocol_filename_requires_no_pdf_parsing(self):
        path = Path(
            "[SBIS-019fb7a8-c6b2][Пиляева] "
            "Уведомление о продаже помещения.pdf"
        )
        parsed = parse_protocol_filename(path)
        self.assertEqual(parsed["sbis_id"], "019fb7a8-c6b2")
        self.assertEqual(derive_surname(path, Path("C:/SBIS")), "Пиляева")
        self.assertEqual(
            content_from_file(path),
            "Уведомление о продаже помещения",
        )

    def test_v2_person_uses_full_fio(self):
        path = Path(
            "[ФЛ][Иванов Иван Иванович] Обращение о перерасчете.pdf"
        )
        parsed = parse_protocol_filename(path)
        self.assertEqual(parsed["correspondent_type"], "person")
        self.assertEqual(parsed["correspondent"], "Иванов Иван Иванович")
        self.assertEqual(parsed["content"], "Обращение о перерасчете")

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            target = root / path.name
            target.write_bytes(b"pdf")
            item = prepare_items(
                root, [target],
                {"version": 1, "next_number": 1, "files": {}},
            )[0]
            self.assertEqual(
                item["doc_data"]["корреспондент"],
                "Иванов Иван Иванович",
            )
            self.assertEqual(item["doc_data"]["корреспондент_тип"], "person")
            self.assertEqual(item["doc_data"]["тип_индекс"], 8)

    def test_v2_legal_uses_full_name_and_legal_kind(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            target = root / "[ЮЛ][Тепловая Компания] Письмо в ОмскРТС.pdf"
            target.write_bytes(b"pdf")
            item = prepare_items(
                root, [target],
                {"version": 1, "next_number": 1, "files": {}},
            )[0]
            self.assertEqual(item["correspondent"], "Тепловая Компания")
            self.assertEqual(
                item["doc_data"]["корреспондент"],
                "Тепловая Компания",
            )
            self.assertEqual(item["doc_data"]["корреспондент_тип"], "legal")
            self.assertEqual(item["doc_data"]["тип_индекс"], 5)

    def test_hidden_manifest_id_prevents_duplicate_after_rename(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            first = root / "[ФЛ][Иванов Иван Иванович] Первое имя.pdf"
            first.write_bytes(b"pdf")
            manifest = {
                "doc-77|pdf-asud-v2": {
                    "sbis_id": "doc-77",
                    "revision": "rev-1",
                    "files": [first.name],
                    "archive": "",
                },
            }
            manifest_path = root / ".manifest.json"
            manifest_path.write_text(
                json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
            state = {"version": 1, "next_number": 1, "files": {}}

            ids = load_manifest_ids(root)
            items = prepare_items(root, [first], state, manifest_ids=ids)
            entry = state["files"][items[0]["key"]]
            entry["status"] = "OK"

            second = root / "[ФЛ][Иванов Иван Иванович] Второе имя.pdf"
            first.rename(second)
            manifest["doc-77|pdf-asud-v2"]["files"] = [second.name]
            manifest_path.write_text(
                json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
            again = prepare_items(
                root, [second], state,
                manifest_ids=load_manifest_ids(root),
            )
            self.assertEqual(again[0]["key"], "sbis:doc-77")
            self.assertTrue(again[0]["already_done"])

    def test_protocol_id_survives_rename_and_prevents_duplicate(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            first = root / "[SBIS-doc-42][Иванов] Первое имя.pdf"
            first.write_bytes(b"pdf")
            state = {"version": 1, "next_number": 1, "files": {}}

            items = prepare_items(root, scan_files(root), state)
            self.assertEqual(items[0]["number"], 1)
            entry = state["files"][items[0]["key"]]
            entry["status"] = "OK"
            entry["signature"] = items[0]["signature"]

            second = root / "[SBIS-doc-42][Иванов] Новое имя.pdf"
            first.rename(second)
            again = prepare_items(root, scan_files(root), state)
            self.assertEqual(again[0]["number"], 1)
            self.assertTrue(again[0]["already_done"])
            self.assertEqual(again[0]["content"], "Новое имя")

    def test_scan_and_stable_numbering(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            first = root / "Иванов Письмо.pdf"
            second = root / "Петров Ответ.docx"
            first.write_bytes(b"pdf")
            second.write_bytes(b"docx")
            (root / ".asud_sbis_state.json").write_text("{}", encoding="utf-8")
            (root / "СБИС_резолюции.xlsx").write_bytes(b"xlsx")

            files = scan_files(root)
            self.assertEqual([p.name for p in files],
                             ["Иванов Письмо.pdf", "Петров Ответ.docx"])

            state = {"version": 1, "next_number": 1, "files": {}}
            items = prepare_items(root, files, state)
            self.assertEqual([item["number"] for item in items], [1, 2])
            self.assertEqual(items[0]["doc_data"]["номер_обращения"], "б/н (1)")
            self.assertEqual(items[0]["doc_data"]["корреспондент"], "Иванов - -")

            first_entry = state["files"][items[0]["key"]]
            first_entry["status"] = "OK"
            first_entry["signature"] = items[0]["signature"]
            again = prepare_items(root, files, state)
            self.assertTrue(again[0]["already_done"])
            self.assertEqual([item["number"] for item in again], [1, 2])


if __name__ == "__main__":
    unittest.main()
