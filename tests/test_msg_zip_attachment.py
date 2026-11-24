import tempfile
import unittest
import zipfile
from email import policy
from email.parser import BytesParser
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from docx import Document

from shared import attachments


class _FakeAttachment:
    def __init__(self, name, data):
        self.name = name
        self.data = data
        self.calls = []

    def save(self, *, customPath, extractEmbedded, zip=None):
        self.calls.append((customPath, extractEmbedded))
        if zip is not None:
            zip.writestr(f"{customPath}/{self.name}", self.data)
            return ("file", f"{customPath}/{self.name}")
        path = Path(customPath) / self.name
        path.write_bytes(self.data)
        return ("file", str(path))


class _FakeMessage:
    subject = "Тема из MSG"
    body = "Тело из MSG"
    sender = "sender@example.test"
    to = "receiver@example.test"
    cc = "copy@example.test"
    date = "10.11.2026 12:00"

    def __init__(self, message_attachments=None):
        self.attachments = message_attachments or []
        self.closed = False

    def close(self):
        self.closed = True


class MsgBundleTests(unittest.TestCase):
    def test_build_msg_zip_contains_full_letter_and_all_attachments(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "письмо.msg"
            target = root / "письмо.zip"
            source.touch()
            item = _FakeAttachment("document.pdf", b"PDF data")
            message = _FakeMessage([item])

            with patch.object(
                    attachments.extract_msg, "openMsg", return_value=message):
                result = attachments.build_msg_zip(
                    source,
                    target,
                    subject="Полная тема после декодирования",
                    body="Полный текст после декодирования",
                )

            self.assertEqual(result, str(target.resolve()))
            self.assertTrue(message.closed)
            self.assertEqual(item.calls, [("Вложения", True)])
            with zipfile.ZipFile(target) as archive:
                self.assertEqual(
                    set(archive.namelist()),
                    {"Письмо.docx", "Вложения/document.pdf"},
                )
                letter = Document(BytesIO(archive.read("Письмо.docx")))
                self.assertEqual(letter.paragraphs[0].text, "Обращение")
                metadata = {
                    row.cells[0].text: row.cells[1].text
                    for row in letter.tables[0].rows
                }
                self.assertEqual(
                    metadata,
                    {
                        "Тема": "Полная тема после декодирования",
                        "От": "sender@example.test",
                        "Кому": "receiver@example.test",
                        "Копия": "copy@example.test",
                        "Дата": "10.11.2026 12:00",
                    },
                )
                paragraph_text = "\n".join(p.text for p in letter.paragraphs)
                self.assertIn("Текст обращения", paragraph_text)
                self.assertIn(
                    "Полный текст после декодирования",
                    paragraph_text,
                )
                self.assertEqual(
                    archive.read("Вложения/document.pdf"), b"PDF data"
                )

    def test_build_msg_eml_contains_docx_and_all_attachments(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "письмо.msg"
            target = root / "письмо.eml"
            source.touch()
            item = _FakeAttachment("document.pdf", b"PDF data")
            message = _FakeMessage([item])

            with patch.object(
                    attachments.extract_msg, "openMsg", return_value=message):
                result = attachments.build_msg_eml(
                    source,
                    target,
                    subject="Полная тема после декодирования",
                    body="Полный текст после декодирования",
                )

            self.assertEqual(result, str(target.resolve()))
            self.assertTrue(message.closed)
            parsed = BytesParser(policy=policy.default).parsebytes(
                target.read_bytes()
            )
            self.assertEqual(
                parsed["Subject"], "Полная тема после декодирования"
            )
            self.assertIn(
                "Полный текст после декодирования",
                parsed.get_body(preferencelist=("plain",)).get_content(),
            )
            parts = {
                part.get_filename(): part
                for part in parsed.iter_attachments()
            }
            self.assertEqual(
                set(parts), {"Обращение.docx", "document.pdf"}
            )
            letter = Document(BytesIO(parts["Обращение.docx"].get_content()))
            self.assertEqual(letter.paragraphs[0].text, "Обращение")
            self.assertIn(
                "Полный текст после декодирования",
                "\n".join(p.text for p in letter.paragraphs),
            )
            self.assertEqual(parts["document.pdf"].get_content(), b"PDF data")

    def test_old_zip_setting_migrates_to_eml_and_keeps_source_msg(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "Тест письмо.msg"
            source.touch()
            seen = {}

            def fake_build(_source, output, **_kwargs):
                Path(output).write_bytes(b"EML")
                return output

            def fake_upload(_driver, upload_path, max_attempts=3):
                seen["path"] = Path(upload_path)
                seen["max_attempts"] = max_attempts
                self.assertTrue(seen["path"].is_file())
                self.assertEqual(seen["path"].suffix, ".eml")
                return True

            with (
                patch.object(attachments, "build_msg_eml", side_effect=fake_build),
                patch.object(attachments, "build_msg_zip") as legacy_builder,
                patch.object(
                    attachments, "_attach_content_file", side_effect=fake_upload
                ),
            ):
                result = attachments.attach_content(
                    object(),
                    source,
                    max_attempts=4,
                    msg_attachment_mode="zip",
                    email_subject="Тема",
                    email_body="Тело",
                )

            self.assertTrue(result)
            self.assertEqual(seen["max_attempts"], 4)
            self.assertFalse(seen["path"].exists())
            self.assertTrue(source.exists())
            legacy_builder.assert_not_called()

    def test_legacy_zip_mode_remains_available_for_diagnostics(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "letter.msg"
            source.touch()

            def fake_build(_source, output, **_kwargs):
                with zipfile.ZipFile(output, "w") as archive:
                    archive.writestr("Обращение.docx", b"docx")
                return output

            with (
                patch.object(attachments, "build_msg_zip", side_effect=fake_build),
                patch.object(
                    attachments, "_attach_content_file", return_value=True
                ) as upload,
            ):
                result = attachments.attach_content(
                    object(), source, msg_attachment_mode="zip-legacy"
                )

            self.assertTrue(result)
            self.assertEqual(Path(upload.call_args.args[1]).suffix, ".zip")

    def test_broken_msg_is_not_uploaded_or_deleted(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "broken.msg"
            source.touch()
            with (
                patch.object(
                    attachments, "build_msg_eml", side_effect=ValueError("broken")
                ),
                patch.object(attachments, "_attach_content_file") as upload,
            ):
                result = attachments.attach_content(
                    object(), source, msg_attachment_mode="zip"
                )

            self.assertFalse(result)
            upload.assert_not_called()
            self.assertTrue(source.exists())

    def test_original_mode_passes_original_path(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "letter.msg"
            source.touch()
            driver = object()
            with patch.object(
                    attachments, "_attach_content_file", return_value=True) as upload:
                result = attachments.attach_content(
                    driver, source, msg_attachment_mode="original"
                )

            self.assertTrue(result)
            upload.assert_called_once_with(
                driver, str(source.resolve()), max_attempts=3
            )


if __name__ == "__main__":
    unittest.main()
