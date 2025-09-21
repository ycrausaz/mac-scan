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

    def _build_job_xml(self, dpi=300, color_mode="Color", page_size="A4", source="Platen"):
        """
        Build scan job XML for single-sided scanning only.
        """
        color_mode = "Grayscale" if color_mode.lower().startswith("gray") else "Color"
        
        xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<scan:ScanSettings xmlns:scan="http://schemas.hp.com/imaging/escl/2011/05/03" xmlns:pwg="http://www.pwg.org/schemas/2010/12/sm">
  <pwg:Version>2.0</pwg:Version>
  <pwg:InputSource>{source}</pwg:InputSource>
  <pwg:DocumentFormat>application/pdf</pwg:DocumentFormat>
  <pwg:ColorMode>{color_mode}</pwg:ColorMode>
  <pwg:Duplex>Simplex</pwg:Duplex>
  <pwg:MediaSizeName>{page_size}</pwg:MediaSizeName>
  <scan:Resolution>
    <scan:XResolution>{dpi}</scan:XResolution>
    <scan:YResolution>{dpi}</scan:YResolution>
  </scan:Resolution>
  <scan:Intent>Document</scan:Intent>
</scan:ScanSettings>"""
        
        return xml.strip()

    def start_job(self, dpi=300, color_mode="Color", page_size="A4", input_source=None):
        """
        Start a single-sided scan job.
        
        Args:
            dpi: Scan resolution (75-1200)
            color_mode: "Color" or "Grayscale"
            page_size: Paper size ("A4", "Letter", "Legal", etc.)
            input_source: Override source selection ("Auto", "Feeder", "Platen", or None for auto-detect)
        
        Returns:
            job_location: URL for the started scan job
        """
        # Determine input source
        if input_source == "Feeder":
            source = "Feeder"
        elif input_source == "Platen":
            source = "Platen"
        else:
            # Auto-detect or legacy behavior
            source = self.choose_input_source()

        url = urljoin(self.base, "ScanJobs")
        xml = self._build_job_xml(dpi, color_mode, page_size, source)
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
            xml2 = self._build_job_xml(dpi, color_mode, page_size, fallback_source)
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
        """
        Fetch the scanned PDF from a started job.
        
        Args:
            job_location: URL returned from start_job()
            out_path: Local file path to save the PDF
            
        Returns:
            out_path: The same path that was provided
        """
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

    def scan_to_pdf(self, out_path, dpi=300, color_mode="Color", page_size="A4", input_source=None):
        """
        Complete single-sided scan operation: start job and fetch PDF.
        
        Args:
            out_path: Local file path to save the PDF
            dpi: Scan resolution (75-1200)
            color_mode: "Color" or "Grayscale"
            page_size: Paper size ("A4", "Letter", "Legal", etc.)
            input_source: Override source selection ("Auto", "Feeder", "Platen", or None for auto-detect)
            
        Returns:
            out_path: The same path that was provided
        """
        job_loc = self.start_job(
            dpi=dpi, 
            color_mode=color_mode, 
            page_size=page_size,
            input_source=input_source
        )
        return self.fetch_pdf(job_loc, out_path)

    def get_status(self):
        """
        Get the current scanner status as XML.
        
        Returns:
            XML string containing scanner status information
        """
        url = urljoin(self.base, "ScannerStatus")
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return r.text

    def choose_input_source(self):
        """
        Inspect status and choose 'Feeder' if ADF has pages; else 'Platen'.
        """
        try:
            xml = self.get_status()
            root = ET.fromstring(xml)
            
            # Check for various ADF status indicators
            text = xml.lower()
            
            # ADF detection patterns
            adf_indicators = [
                ("adf" in text and "loaded" in text),
                ("adf" in text and "present" in text), 
                ("adf" in text and "haspaper" in text),
                ("feeder" in text and "loaded" in text),
                ("inputsource" in text and "feeder" in text),
                ("media" in text and "feeder" in text),
                ("documentfeeder" in text and ("loaded" in text or "ready" in text)),
                ("inputtray" in text and "adf" in text),
                ("paper" in text and ("adf" in text or "feeder" in text)),
            ]
            
            adf_has_pages = any(adf_indicators)
            
            # Try to parse XML structure for more reliable detection
            try:
                for elem in root.iter():
                    tag_lower = elem.tag.lower()
                    text_lower = (elem.text or "").lower()
                    
                    if "input" in tag_lower and "source" in tag_lower:
                        if "feeder" in text_lower or "adf" in text_lower:
                            adf_has_pages = True
                            break
                    
                    if any(keyword in tag_lower for keyword in ["media", "paper", "document"]):
                        if any(status in text_lower for status in ["loaded", "present", "ready"]):
                            if any(source in tag_lower for source in ["adf", "feeder"]):
                                adf_has_pages = True
                                break
            except Exception:
                pass
            
            if adf_has_pages:
                return "Feeder"
                
        except Exception:
            pass
            
        return "Platen"

    def get_scanner_capabilities(self):
        """
        Get scanner capabilities (optional, for advanced usage).
        
        Returns:
            XML string containing scanner capabilities, or None if not supported
        """
        try:
            url = urljoin(self.base, "ScannerCapabilities")
            r = requests.get(url, timeout=10)
            r.raise_for_status()
            return r.text
        except Exception:
            return None

    def cancel_job(self, job_location):
        """
        Cancel a running scan job.
        
        Args:
            job_location: URL of the job to cancel
            
        Returns:
            True if successful, False otherwise
        """
        try:
            r = requests.delete(job_location, timeout=5)
            return 200 <= r.status_code < 300
        except Exception:
            return False

    def list_jobs(self):
        """
        List active scan jobs (optional, for advanced usage).
        
        Returns:
            XML string containing active jobs, or None if not supported
        """
        try:
            url = urljoin(self.base, "ScanJobs")
            r = requests.get(url, timeout=10)
            r.raise_for_status()
            return r.text
        except Exception:
            return None

    def test_connection(self):
        """
        Test if the scanner is reachable and responding.
        
        Returns:
            tuple: (is_reachable: bool, status_code: int, error_message: str or None)
        """
        try:
            url = urljoin(self.base, "ScannerStatus")
            r = requests.get(url, timeout=5)
            return (True, r.status_code, None)
        except requests.exceptions.Timeout:
            return (False, 0, "Connection timeout")
        except requests.exceptions.ConnectionError:
            return (False, 0, "Connection refused or network unreachable")
        except requests.exceptions.RequestException as e:
            return (False, 0, str(e))
        except Exception as e:
            return (False, 0, f"Unexpected error: {e}")

    def debug_scan_settings(self, dpi=300, color_mode="Color", page_size="A4", input_source=None):
        """
        Generate scan settings XML for debugging purposes without starting a job.
        
        Returns:
            tuple: (source: str, xml: str)
        """
        # Determine input source (same logic as start_job)
        if input_source == "Feeder":
            source = "Feeder"
        elif input_source == "Platen":
            source = "Platen"
        else:
            # Auto-detect
            source = self.choose_input_source()

        xml = self._build_job_xml(dpi, color_mode, page_size, source)
        return (source, xml)
