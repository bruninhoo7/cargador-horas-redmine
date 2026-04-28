"""
Verificación y descarga de actualizaciones desde GitHub Releases.
"""
import json
import os, urllib.request
import ssl

GH_USER = "bruninhoo7"
GH_REPO = "cargador-horas-redmine"


def verificar_actualizacion(app_version):
    """Consulta GitHub para ver si hay una versión más reciente.

    Retorna (tag, url_exe) si hay actualización, o (None, None) si no.
    Hace varios reintentos con timeouts crecientes y fallback HTTP/SSL.
    """
    urls = [
        f"https://api.github.com/repos/{GH_USER}/{GH_REPO}/releases/latest",
        f"http://api.github.com/repos/{GH_USER}/{GH_REPO}/releases/latest",
    ]

    def v2t(v):
        try:
            return tuple(int(x) for x in v.split("."))
        except Exception:
            return (0,)

    for url in urls:
        for timeout in [8, 15, 30]:
            for ctx in [None, ssl._create_unverified_context()]:
                try:
                    req = urllib.request.Request(url)
                    req.add_header("Accept", "application/vnd.github.v3+json")
                    req.add_header("User-Agent", "CargadorHorasRedmine")
                    kwargs = {"timeout": timeout}
                    if ctx is not None and url.startswith("https"):
                        kwargs["context"] = ctx
                    with urllib.request.urlopen(req, **kwargs) as resp:
                        data = json.loads(resp.read())
                    latest = data.get("tag_name", "").lstrip("v").strip()
                    if latest and v2t(latest) > v2t(app_version):
                        assets = data.get("assets", [])
                        url_exe = next(
                            (a["browser_download_url"] for a in assets
                             if a["name"].endswith(".exe")),
                            None
                        )
                        if url_exe:
                            return latest, url_exe
                    return None, None
                except Exception:
                    continue
    return None, None


def descargar_actualizacion(url_asset, dest_path, progress_cb=None):
    """Descarga el exe de la nueva versión, llamando a progress_cb(porcentaje).
    Reintenta con SSL desactivado y HTTP si falla."""
    # Generar URLs alternativas (https con SSL, https sin SSL, http)
    urls_to_try = [url_asset]
    if url_asset.startswith("https://"):
        urls_to_try.append(url_asset.replace("https://", "http://", 1))

    for url in urls_to_try:
        for ctx in [None, ssl._create_unverified_context()]:
            try:
                req = urllib.request.Request(url)
                req.add_header("User-Agent", "CargadorHorasRedmine")
                kwargs = {"timeout": 120}
                if ctx is not None and url.startswith("https"):
                    kwargs["context"] = ctx
                with urllib.request.urlopen(req, **kwargs) as resp:
                    total = int(resp.headers.get("Content-Length", 0))
                    downloaded = 0
                    with open(dest_path, "wb") as f:
                        while True:
                            chunk = resp.read(8192)
                            if not chunk:
                                break
                            f.write(chunk)
                            downloaded += len(chunk)
                            if progress_cb and total:
                                progress_cb(downloaded / total * 100)
                # Verificar que se descargo algo significativo
                if os.path.getsize(dest_path) < 1000:
                    continue  # Probablemente fallo, intentar siguiente
                return True
            except Exception:
                continue
    return False
