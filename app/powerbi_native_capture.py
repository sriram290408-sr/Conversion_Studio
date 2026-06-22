"""Power BI Service renderer for Power BI-native visuals.

This module implements a no-placeholder policy. Native/custom Power BI visuals
are exported by Power BI Service, cropped with PBIX coordinates, and embedded in
Excel. If the service is not configured, the build fails rather than silently
creating a placeholder.
"""

from __future__ import annotations

import io
import json
import logging
import os
import time
import zipfile
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import requests
from PIL import Image

logger = logging.getLogger("powerbi_native_capture")


class PowerBINativeCaptureError(RuntimeError):
    pass


class PowerBINativeCaptureService:
    def __init__(self) -> None:
        self.tenant_id = os.getenv("PBI_TENANT_ID", "").strip()
        self.client_id = os.getenv("PBI_CLIENT_ID", "").strip()
        self.client_secret = os.getenv("PBI_CLIENT_SECRET", "").strip()
        self.workspace_id = os.getenv("PBI_WORKSPACE_ID", "").strip()
        self.report_id = os.getenv("PBI_REPORT_ID", "").strip()
        self.dataset_id = os.getenv("PBI_DATASET_ID", "").strip()
        self.scope = "https://analysis.windows.net/powerbi/api/.default"
        self.api = "https://api.powerbi.com/v1.0/myorg"
        self.timeout = int(os.getenv("PBI_HTTP_TIMEOUT_SECONDS", "90"))
        self.poll_seconds = float(os.getenv("PBI_EXPORT_POLL_SECONDS", "2"))
        self.max_wait = int(os.getenv("PBI_EXPORT_MAX_WAIT_SECONDS", "300"))
        self._token: Optional[str] = None
        self._page_cache: Dict[str, Path] = {}

    @property
    def configured(self) -> bool:
        return all(
            (
                self.tenant_id,
                self.client_id,
                self.client_secret,
                self.workspace_id,
                self.report_id,
            )
        )

    def require_configured(self) -> None:
        if self.configured:
            return
        missing = [
            name
            for name, value in {
                "PBI_TENANT_ID": self.tenant_id,
                "PBI_CLIENT_ID": self.client_id,
                "PBI_CLIENT_SECRET": self.client_secret,
                "PBI_WORKSPACE_ID": self.workspace_id,
                "PBI_REPORT_ID": self.report_id,
            }.items()
            if not value
        ]
        raise PowerBINativeCaptureError(
            "Power BI-native rendering requires Power BI Service configuration. "
            f"Missing environment variables: {', '.join(missing)}"
        )

    def _access_token(self) -> str:
        self.require_configured()
        if self._token:
            return self._token
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        response = requests.post(
            url,
            data={
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "grant_type": "client_credentials",
                "scope": self.scope,
            },
            timeout=self.timeout,
        )
        if not response.ok:
            raise PowerBINativeCaptureError(
                f"Power BI authentication failed ({response.status_code}): {response.text[:500]}"
            )
        self._token = str(response.json().get("access_token") or "")
        if not self._token:
            raise PowerBINativeCaptureError(
                "Power BI token response did not contain access_token"
            )
        return self._token

    def _headers(self, json_body: bool = True) -> Dict[str, str]:
        headers = {"Authorization": f"Bearer {self._access_token()}"}
        if json_body:
            headers["Content-Type"] = "application/json"
        return headers

    def refresh_dataset(self) -> None:
        if not self.dataset_id:
            return
        url = f"{self.api}/groups/{self.workspace_id}/datasets/{self.dataset_id}/refreshes"
        response = requests.post(
            url,
            headers=self._headers(),
            json={"notifyOption": "NoNotification"},
            timeout=self.timeout,
        )
        if response.status_code not in (200, 202):
            raise PowerBINativeCaptureError(
                f"Dataset refresh could not be started ({response.status_code}): {response.text[:500]}"
            )
        deadline = time.time() + self.max_wait
        while time.time() < deadline:
            history = requests.get(
                url + "?$top=1", headers=self._headers(False), timeout=self.timeout
            )
            if history.ok:
                values = history.json().get("value") or []
                status = str(values[0].get("status") if values else "").casefold()
                if status in {"completed", "success"}:
                    return
                if status in {"failed", "disabled", "cancelled"}:
                    raise PowerBINativeCaptureError(
                        f"Dataset refresh ended with status {status}"
                    )
            time.sleep(self.poll_seconds)
        raise PowerBINativeCaptureError(
            "Timed out waiting for Power BI dataset refresh"
        )

    def export_page_png(
        self, page_name: str, output_dir: Path, filters: Optional[list] = None
    ) -> Path:
        self.require_configured()
        cache_key = json.dumps(
            {"page": page_name, "filters": filters or []}, sort_keys=True
        )
        if cache_key in self._page_cache and self._page_cache[cache_key].exists():
            return self._page_cache[cache_key]

        output_dir.mkdir(parents=True, exist_ok=True)
        payload: Dict[str, Any] = {
            "format": "PNG",
            "powerBIReportConfiguration": {"pages": [{"pageName": page_name}]},
        }
        if filters:
            payload["powerBIReportConfiguration"]["reportLevelFilters"] = filters

        start_url = (
            f"{self.api}/groups/{self.workspace_id}/reports/{self.report_id}/ExportTo"
        )
        response = requests.post(
            start_url, headers=self._headers(), json=payload, timeout=self.timeout
        )
        if response.status_code not in (200, 202):
            raise PowerBINativeCaptureError(
                f"Power BI PNG export failed to start ({response.status_code}): {response.text[:800]}"
            )
        export_id = str(response.json().get("id") or "")
        if not export_id:
            raise PowerBINativeCaptureError(
                "Power BI export response did not contain an export id"
            )

        status_url = f"{self.api}/groups/{self.workspace_id}/reports/{self.report_id}/exports/{export_id}"
        deadline = time.time() + self.max_wait
        while time.time() < deadline:
            status_response = requests.get(
                status_url, headers=self._headers(False), timeout=self.timeout
            )
            if not status_response.ok:
                raise PowerBINativeCaptureError(
                    f"Power BI export status failed ({status_response.status_code}): {status_response.text[:500]}"
                )
            status = str(status_response.json().get("status") or "").casefold()
            if status == "succeeded":
                break
            if status in {"failed", "cancelled"}:
                raise PowerBINativeCaptureError(
                    f"Power BI export ended with status {status}"
                )
            time.sleep(self.poll_seconds)
        else:
            raise PowerBINativeCaptureError("Timed out waiting for Power BI PNG export")

        file_response = requests.get(
            status_url + "/file", headers=self._headers(False), timeout=self.timeout
        )
        if not file_response.ok:
            raise PowerBINativeCaptureError(
                f"Power BI export download failed ({file_response.status_code}): {file_response.text[:500]}"
            )

        content = file_response.content
        page_file = output_dir / f"powerbi_page_{self._safe_name(page_name)}.png"
        if content[:2] == b"PK":
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                png_names = [
                    n for n in archive.namelist() if n.lower().endswith(".png")
                ]
                if not png_names:
                    raise PowerBINativeCaptureError(
                        "Power BI PNG export ZIP contained no PNG pages"
                    )
                selected = next(
                    (
                        n
                        for n in png_names
                        if page_name.casefold() in Path(n).stem.casefold()
                    ),
                    png_names[0],
                )
                page_file.write_bytes(archive.read(selected))
        else:
            page_file.write_bytes(content)

        Image.open(page_file).verify()
        self._page_cache[cache_key] = page_file
        logger.info("Power BI page rendered: page=%s path=%s", page_name, page_file)
        return page_file

    def crop_visual(
        self,
        page_image: Path,
        output_path: Path,
        layout: Dict[str, Any],
        canvas: Dict[str, Any],
    ) -> Path:
        image = Image.open(page_image).convert("RGBA")
        page_width = max(float(canvas.get("width") or 1280), 1.0)
        page_height = max(float(canvas.get("height") or 720), 1.0)
        sx, sy = image.width / page_width, image.height / page_height
        x = float(layout.get("x") or 0)
        y = float(layout.get("y") or 0)
        w = float(layout.get("width") or 1)
        h = float(layout.get("height") or 1)
        left = max(0, round(x * sx))
        top = max(0, round(y * sy))
        right = min(image.width, round((x + w) * sx))
        bottom = min(image.height, round((y + h) * sy))
        if right <= left or bottom <= top:
            raise PowerBINativeCaptureError(
                f"Invalid Power BI visual crop: {(left, top, right, bottom)}"
            )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        image.crop((left, top, right, bottom)).save(output_path)
        return output_path

    def render_visual(
        self, binding: Dict[str, Any], canvas: Dict[str, Any], output_dir: Path
    ) -> Path:
        page_name = str(
            binding.get("powerbi_page_name")
            or binding.get("page_internal_name")
            or binding.get("page_name")
            or ""
        )
        if not page_name:
            raise PowerBINativeCaptureError(
                "Power BI-native visual is missing page name"
            )
        visual_id = str(
            binding.get("visual_id") or binding.get("chunk_id") or "native_visual"
        )
        page_image = self.export_page_png(
            page_name, output_dir, binding.get("powerbi_filters")
        )
        return self.crop_visual(
            page_image,
            output_dir / f"pbi_native_{self._safe_name(visual_id)}.png",
            dict(binding.get("layout") or {}),
            canvas,
        )

    @staticmethod
    def _safe_name(value: str) -> str:
        return "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in str(value))[
            :100
        ]
