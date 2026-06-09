from __future__ import annotations

from pathlib import Path
import unittest

try:
    from fastapi.testclient import TestClient
    from backend.app.main import app
except ModuleNotFoundError as exc:  # pragma: no cover - 未安装 Web 依赖时让 unittest 明确跳过。
    raise unittest.SkipTest(f"缺少 Web 测试依赖：{exc.name}") from exc

from scripts.build_qingcha_table import required_source_filenames

BASE_DIR = Path(__file__).resolve().parents[1]
COMPLETE_DIR = BASE_DIR / "完整"


def upload_payload(filenames: list[str]) -> list[tuple[str, tuple[str, object, str]]]:
    payload = []
    for filename in filenames:
        payload.append((
            "files",
            (
                filename,
                (COMPLETE_DIR / filename).open("rb"),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ),
        ))
    return payload


def close_payload(payload: list[tuple[str, tuple[str, object, str]]]) -> None:
    for _, (_, file_obj, _) in payload:
        file_obj.close()


class WebApiTest(unittest.TestCase):
    def setUp(self) -> None:
        self.client = TestClient(app)

    def test_required_files_excludes_template(self) -> None:
        response = self.client.get("/api/required-files")

        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertIs(data["templateProvidedByServer"], True)
        self.assertEqual(data["templateName"], "清查表.xlsx")
        self.assertEqual(data["requiredFiles"], required_source_filenames())
        self.assertNotIn("清查表.xlsx", data["requiredFiles"])

    def test_missing_upload_returns_logs_and_missing_files(self) -> None:
        filenames = [filename for filename in required_source_filenames() if filename != "合同明细数据.xlsx"]
        payload = upload_payload(filenames)
        try:
            response = self.client.post("/api/build", files=payload)
        finally:
            close_payload(payload)

        self.assertEqual(response.status_code, 400)
        data = response.json()
        self.assertEqual(data["error"], "missing_required_files")
        self.assertEqual(data["missingFiles"], ["合同明细数据.xlsx"])
        self.assertIn("WARN: 缺少上传文件：合同明细数据.xlsx", data["logs"])

    def test_upload_accepts_required_filename_prefix(self) -> None:
        payload = upload_payload(required_source_filenames())
        prefixed_payload = []
        renamed_file = None
        for field, (filename, file_obj, content_type) in payload:
            if filename == "固定资产统计数据.xlsx":
                renamed_file = "固定资产统计数据-南街村.xlsx"
                prefixed_payload.append((field, (renamed_file, file_obj, content_type)))
            else:
                prefixed_payload.append((field, (filename, file_obj, content_type)))

        try:
            response = self.client.post("/api/build", files=prefixed_payload)
        finally:
            close_payload(payload)

        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertIn(f"WARN: 已将上传文件 {renamed_file} 识别为 固定资产统计数据.xlsx", data["logs"])

    def test_custom_output_filename_is_used_for_download(self) -> None:
        payload = upload_payload(required_source_filenames())
        try:
            response = self.client.post(
                "/api/build",
                data={"output_filename": "南街村-清查表（已填充）.xlsx"},
                files=payload,
            )
        finally:
            close_payload(payload)

        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["filename"], "南街村-清查表（已填充）.xlsx")

        download = self.client.get(data["downloadUrl"])
        self.assertEqual(download.status_code, 200)
        self.assertIn("filename*=utf-8''", download.headers["content-disposition"])
        self.assertIn("%E5%8D%97%E8%A1%97%E6%9D%91-%E6%B8%85%E6%9F%A5%E8%A1%A8", download.headers["content-disposition"])

    def test_successful_upload_returns_downloadable_workbook(self) -> None:
        payload = upload_payload(required_source_filenames())
        try:
            response = self.client.post("/api/build", files=payload)
        finally:
            close_payload(payload)

        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["filename"], "清查表（已填充）.xlsx")
        self.assertTrue(data["downloadUrl"].startswith("/api/download/"))
        self.assertTrue(any(log.startswith("INFO: 已生成工作簿：") for log in data["logs"]))

        download = self.client.get(data["downloadUrl"])
        self.assertEqual(download.status_code, 200)
        self.assertEqual(download.headers["content-type"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.assertTrue(download.content.startswith(b"PK"))
        self.assertGreater(len(download.content), 10000)


if __name__ == "__main__":
    unittest.main()
