import requests
from urllib.parse import urljoin
import time
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

class ESCLScanner:
    def __init__(self, host: str):
        # host: "192.168.1.6"
        host = host.strip()
        base = host if host.startswith("http") else f"http://{host}"
        # eSCL base path
        if not base.endswith("/"):
            base += "/"
        self.base = urljoin(base, "eSCL/")

    def _build_job_xml(self, dpi=300, color_mode="Color", duplex=False, page_size="A4", source="Platen"):
        duplex_str = "Duplex" if duplex and source == "Feeder" else "Simplex"
        color_mode = "Grayscale" if color_mode.lower().startswith("gray") else "Color"
        return f"""<?xml version="1.0" encoding="UTF-8"?>
<scan:ScanSettings xmlns:scan="http://schemas.hp.com/imaging/escl/2011/05/03" xmlns:pwg="http://www.pwg.org/schemas/2010/12/sm">
  <pwg:Version>2.0</pwg:Version>
  <pwg:InputSource>{source}</pwg:InputSource>
  <pwg:DocumentFormat>application/pdf</pwg:DocumentFormat>
  <pwg:ColorMode>{color_mode}</pwg:ColorMode>
  <pwg:Duplex>{duplex_str}</pwg:Duplex>
  <pwg:MediaSizeName>{page_size}</pwg:MediaSizeName>
  <scan:Resolution>
    <scan:XResolution>{dpi}</scan:XResolution>
    <scan:YResolution>{dpi}</scan:YResolution>
  </scan:Resolution>
  <scan:Intent>Document</scan:Intent>
</scan:ScanSettings>""".strip()

    def start_job(self, dpi=300, color_mode="Color", duplex=False, page_size="A4"):
        # Pick a sane source based on device state
        source = self.choose_input_source()
        if source == "Platen":
            duplex = False  # platen can't duplex

        url = urljoin(self.base, "ScanJobs")
        xml = self._build_job_xml(dpi, color_mode, duplex, page_size, source)
        headers_main = {
            "Content-Type": "text/xml; charset=utf-8",
            "Accept": "application/pdf",
            "Connection": "keep-alive",
        }

        # Try a few times if we get 409 (device busy / previous job not cleared)
        last_err = None
        for attempt in range(6):  # ~6 seconds total
            try:
                r = requests.post(url, data=xml.encode("utf-8"), headers=headers_main, timeout=30)
                if r.status_code == 409:
                    # Busy: brief pause then retry
                    time.sleep(1.0)
                    continue
                r.raise_for_status()
                job_location = r.headers.get("Location") or r.headers.get("location")
                if not job_location:
                    raise RuntimeError("Scanner did not provide job Location header")
                if job_location.startswith("/"):
                    job_location = urljoin(self.base, job_location.lstrip("/"))
                return job_location
            except requests.HTTPError as e:
                last_err = e
                # For any other 4xx/5xx, break early
                if getattr(e.response, "status_code", None) != 409:
                    break

        # As a fallback, try forcing the other source once (in case Auto/Platen choice was wrong)
        try:
            fallback_source = "Feeder" if source == "Platen" else "Platen"
            if fallback_source == "Platen":
                duplex = False
            xml2 = self._build_job_xml(dpi, color_mode, duplex, page_size, fallback_source)
            r2 = requests.post(url, data=xml2.encode("utf-8"), headers=headers_main, timeout=30)
            r2.raise_for_status()
            job_location = r2.headers.get("Location") or r2.headers.get("location")
            if not job_location:
                raise RuntimeError("Scanner did not provide job Location header (fallback)")
            if job_location.startswith("/"):
                job_location = urljoin(self.base, job_location.lstrip("/"))
            return job_location
        except Exception:
            pass

        raise RuntimeError(f"Unable to start scan job (last error: {last_err})")

    def fetch_pdf(self, job_location, out_path):
        next_doc = job_location.rstrip("/") + "/NextDocument"
        headers = {"Accept": "application/pdf"}
        try:
            with requests.get(next_doc, headers=headers, stream=True, timeout=120) as r:
                r.raise_for_status()
                with open(out_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=65536):
                        if chunk:
                            f.write(chunk)
            return out_path
        finally:
            # Best-effort cleanup; ignore errors
            try:
                requests.delete(job_location, timeout=5)
            except Exception:
                pass

    def scan_to_pdf(self, out_path, dpi=300, color_mode="Color", duplex=False, page_size="A4"):
        job_loc = self.start_job(dpi=dpi, color_mode=color_mode, duplex=duplex, page_size=page_size)
        return self.fetch_pdf(job_loc, out_path)

    def get_status(self):
        url = urljoin(self.base, "ScannerStatus")
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return r.text

    def choose_input_source(self):
        """
        Inspect status and choose 'Feeder' if ADF has pages; else 'Platen'.
        Many HPs dislike 'Auto' when duplex=true or ADF empty.
        """
        try:
            xml = self.get_status()
            root = ET.fromstring(xml)
            # Different vendors expose slightly different tags; we try a few.
            text = xml.lower()
            adf_has_pages = ("adf" in text and ("loaded" in text or "present" in text or "haspaper" in text))
            if adf_has_pages:
                return "Feeder"
        except Exception:
            pass
        return "Platen"
