"""LibreOffice UNO PDF conversion helpers."""

from __future__ import annotations

import importlib
import shutil
import socket
import subprocess
import tempfile
import time
from dataclasses import dataclass
from io import BufferedWriter
from pathlib import Path
from typing import Any, Protocol

from docxrender.contracts import (
    DocxFieldRefreshOptions,
    DocxToPdfOptions,
    DocxToPdfResult,
)
from docxrender.docx.fields import (
    write_docx_field_update_markers,
    write_frozen_docx_fields,
)

LISTENER_HOST = "127.0.0.1"
LISTENER_START_TIMEOUT_SECONDS = 15.0
LISTENER_POLL_INTERVAL_SECONDS = 0.1
DOCUMENT_LOAD_TIMEOUT_SECONDS = 10.0
DOCUMENT_LOAD_POLL_INTERVAL_SECONDS = 0.2
URL_LIBREOFFICE_PARAMETERS = (
    "https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html"
)
URL_LIBREOFFICE_API = "https://api.libreoffice.org/"
URL_DEBIAN_PYTHON_UNO = "https://packages.debian.org/bullseye/python3-uno"


class UnoUpdatable(Protocol):
    def update(self) -> None: ...


class UnoDocumentIndexes(Protocol):
    def getCount(self) -> int: ...

    def getByIndex(self, index: int) -> UnoUpdatable: ...


class UnoTextFields(Protocol):
    def refresh(self) -> None: ...


class UnoDisposable(Protocol):
    def dispose(self) -> None: ...


class UnoTextDocument(UnoDisposable, Protocol):
    def updateLinks(self) -> None: ...

    def refresh(self) -> None: ...

    def getDocumentIndexes(self) -> UnoDocumentIndexes: ...

    def getTextFields(self) -> UnoTextFields: ...

    def store(self) -> None: ...

    def storeToURL(self, url: str, properties: tuple[Any, ...]) -> None: ...

    def close(self, deliver_ownership: bool) -> None: ...


class UnoDesktop(Protocol):
    def loadComponentFromURL(
        self,
        url: str,
        target_frame_name: str,
        search_flags: int,
        properties: tuple[Any, ...],
    ) -> UnoTextDocument | None: ...


class ListenerProcess(Protocol):
    def poll(self) -> int | None: ...

    def terminate(self) -> None: ...

    def wait(self, timeout: float | None = None) -> int | None: ...

    def kill(self) -> None: ...


@dataclass(frozen=True, slots=True)
class DocxToPdfState:
    options: DocxToPdfOptions


def run_docx_to_pdf_pipeline(options: DocxToPdfOptions) -> DocxToPdfResult:
    state = create_docx_to_pdf_state(options)
    validate_docx_input(state)
    convert_docx_to_pdf_with_uno(state)
    return create_docx_to_pdf_result(state)


def create_docx_to_pdf_state(options: DocxToPdfOptions) -> DocxToPdfState:
    return DocxToPdfState(options=options)


def validate_docx_input(state: DocxToPdfState) -> DocxToPdfState:
    file_in_docx = state.options.file_in_docx
    if not file_in_docx.exists():
        raise FileNotFoundError(f"Input DOCX does not exist: {file_in_docx.resolve()}")
    if not file_in_docx.is_file():
        raise RuntimeError(
            f"Input DOCX is not a regular file: {file_in_docx.resolve()}"
        )
    if file_in_docx.stat().st_size == 0:
        raise RuntimeError(f"Input DOCX is empty: {file_in_docx.resolve()}")
    with file_in_docx.open("rb"):
        pass
    return state


def create_docx_to_pdf_result(state: DocxToPdfState) -> DocxToPdfResult:
    return DocxToPdfResult(
        file_pdf=state.options.file_out_pdf,
        file_docx_refreshed=state.options.file_out_docx_refreshed,
    )


def create_libreoffice_listener_command(
    *,
    exe_libreoffice: Path,
    dir_user_profile: Path,
    port: int,
) -> list[str]:
    return [
        str(exe_libreoffice),
        "--headless",
        f"--accept=socket,host={LISTENER_HOST},port={port};urp;",
        "--norestore",
        "--nodefault",
        f"-env:UserInstallation={dir_user_profile.resolve().as_uri()}",
    ]


def convert_docx_to_pdf_with_uno(state: DocxToPdfState) -> DocxToPdfState:
    options = state.options
    options.dir_user_profile.mkdir(parents=True, exist_ok=True)
    options.file_out_pdf.parent.mkdir(parents=True, exist_ok=True)
    if options.file_out_docx_refreshed is not None:
        options.file_out_docx_refreshed.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="docxrender-docx-stage-") as dir_stage_tmp:
        file_in_docx_staged = copy_docx_to_stage(
            options.file_in_docx,
            dir_stage=Path(dir_stage_tmp),
        )
        if options.should_update_fields:
            write_docx_field_update_markers(file_in_docx_staged)
        uno_module = import_uno_module()
        port = select_free_port()
        (
            file_listener_log_resolved,
            handle_listener_log,
            stdout_listener,
            stderr_listener,
        ) = open_listener_log_handle(options.file_listener_log)
        process_listener = start_libreoffice_listener(
            exe_libreoffice=options.exe_libreoffice,
            dir_user_profile=options.dir_user_profile,
            port=port,
            stdout=stdout_listener,
            stderr=stderr_listener,
            file_listener_log=file_listener_log_resolved,
        )
        try:
            wait_for_listener(port, file_listener_log=file_listener_log_resolved)
            desktop = connect_desktop(uno_module, port)
            doc: UnoTextDocument | None = None
            try:
                doc = load_uno_document_or_raise(
                    uno_module=uno_module,
                    desktop=desktop,
                    file_in_docx_source=options.file_in_docx,
                    file_in_docx_staged=file_in_docx_staged,
                    exe_libreoffice=options.exe_libreoffice,
                    dir_user_profile=options.dir_user_profile,
                    process_listener=process_listener,
                    file_listener_log=file_listener_log_resolved,
                    file_source_lock=find_source_lock_file(options.file_in_docx),
                )
                refresh_uno_document_fields(doc)
                doc.store()
                doc.storeToURL(
                    uno_module.systemPathToFileUrl(str(options.file_out_pdf.resolve())),
                    (
                        create_property("FilterName", "writer_pdf_Export"),
                        create_property("Overwrite", True),
                    ),
                )
                if options.file_out_docx_refreshed is not None:
                    if options.should_freeze_fields:
                        write_frozen_docx_fields(file_in_docx_staged)
                    shutil.copy2(file_in_docx_staged, options.file_out_docx_refreshed)
            finally:
                close_document(doc)
        finally:
            terminate_process(process_listener)
            if handle_listener_log is not None:
                handle_listener_log.close()
    return state


def refresh_docx_with_uno(
    *,
    file_in_docx: Path,
    file_out_docx: Path,
    options: DocxFieldRefreshOptions,
) -> None:
    options.dir_user_profile.mkdir(parents=True, exist_ok=True)
    file_out_docx.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(
        prefix="docxrender-docx-refresh-stage-"
    ) as dir_tmp:
        file_in_docx_staged = copy_docx_to_stage(file_in_docx, dir_stage=Path(dir_tmp))
        uno_module = import_uno_module()
        port = select_free_port()
        (
            file_listener_log_resolved,
            handle_listener_log,
            stdout_listener,
            stderr_listener,
        ) = open_listener_log_handle(options.file_listener_log)
        process_listener = start_libreoffice_listener(
            exe_libreoffice=options.exe_libreoffice,
            dir_user_profile=options.dir_user_profile,
            port=port,
            stdout=stdout_listener,
            stderr=stderr_listener,
            file_listener_log=file_listener_log_resolved,
        )
        try:
            wait_for_listener(port, file_listener_log=file_listener_log_resolved)
            desktop = connect_desktop(uno_module, port)
            doc: UnoTextDocument | None = None
            try:
                doc = load_uno_document_or_raise(
                    uno_module=uno_module,
                    desktop=desktop,
                    file_in_docx_source=file_in_docx,
                    file_in_docx_staged=file_in_docx_staged,
                    exe_libreoffice=options.exe_libreoffice,
                    dir_user_profile=options.dir_user_profile,
                    process_listener=process_listener,
                    file_listener_log=file_listener_log_resolved,
                    file_source_lock=find_source_lock_file(file_in_docx),
                )
                refresh_uno_document_fields(doc)
                doc.store()
            finally:
                close_document(doc)
        finally:
            terminate_process(process_listener)
            if handle_listener_log is not None:
                handle_listener_log.close()
        shutil.copy2(file_in_docx_staged, file_out_docx)


def import_uno_module() -> Any:
    try:
        return importlib.import_module("uno")
    except ImportError as exc:
        raise RuntimeError(
            "\n".join(
                [
                    "error_code=libreoffice_uno_import_failed",
                    "reason=UNO Python bindings are not importable in this Python "
                    "environment.",
                    *create_libreoffice_runtime_guidance_fields(),
                ]
            )
        ) from exc


def create_libreoffice_runtime_guidance_fields() -> list[str]:
    return [
        "runtime_dependency=LibreOffice and Python-UNO are external runtime "
        "dependencies; docxrender does not install them through a Python "
        "package extra.",
        "validate_libreoffice=libreoffice --headless --version",
        'validate_uno=python -c "import uno"',
        "install_debian_ubuntu=sudo apt install libreoffice python3-uno",
        f"docs_libreoffice_parameters={URL_LIBREOFFICE_PARAMETERS}",
        f"docs_libreoffice_api={URL_LIBREOFFICE_API}",
        f"docs_debian_python_uno={URL_DEBIAN_PYTHON_UNO}",
    ]


def validate_libreoffice_executable(exe_libreoffice: Path) -> None:
    if not exe_libreoffice.exists():
        raise FileNotFoundError(
            "\n".join(
                [
                    "error_code=libreoffice_executable_missing",
                    f"exe_libreoffice={exe_libreoffice.resolve()}",
                    *create_libreoffice_runtime_guidance_fields(),
                ]
            )
        )
    if not exe_libreoffice.is_file():
        raise RuntimeError(
            "\n".join(
                [
                    "error_code=libreoffice_executable_not_file",
                    f"exe_libreoffice={exe_libreoffice.resolve()}",
                    *create_libreoffice_runtime_guidance_fields(),
                ]
            )
        )


def start_libreoffice_listener(
    *,
    exe_libreoffice: Path,
    dir_user_profile: Path,
    port: int,
    stdout: BufferedWriter | None,
    stderr: BufferedWriter | None,
    file_listener_log: Path | None,
) -> subprocess.Popen[bytes]:
    validate_libreoffice_executable(exe_libreoffice)
    command = create_libreoffice_listener_command(
        exe_libreoffice=exe_libreoffice,
        dir_user_profile=dir_user_profile,
        port=port,
    )
    try:
        return subprocess.Popen(command, stdout=stdout, stderr=stderr)
    except (FileNotFoundError, PermissionError) as exc:
        raise RuntimeError(
            "\n".join(
                [
                    "error_code=libreoffice_listener_start_failed",
                    f"exe_libreoffice={exe_libreoffice.resolve()}",
                    f"dir_user_profile={dir_user_profile.resolve()}",
                    f"listener_log={listener_log_label(file_listener_log)}",
                    f"launch_error={type(exc).__name__}: {exc}",
                    *create_libreoffice_runtime_guidance_fields(),
                ]
            )
        ) from exc


def select_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind((LISTENER_HOST, 0))
        return int(sock.getsockname()[1])


def wait_for_listener(port: int, *, file_listener_log: Path | None = None) -> None:
    deadline = time.monotonic() + LISTENER_START_TIMEOUT_SECONDS
    while time.monotonic() < deadline:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(LISTENER_POLL_INTERVAL_SECONDS)
            if sock.connect_ex((LISTENER_HOST, port)) == 0:
                return
        time.sleep(LISTENER_POLL_INTERVAL_SECONDS)
    fields = [
        "error_code=libreoffice_uno_listener_timeout",
        f"listener_host={LISTENER_HOST}",
        f"listener_port={port}",
        f"listener_log={listener_log_label(file_listener_log)}",
    ]
    text_log_tail = read_log_tail(file_listener_log)
    if text_log_tail:
        fields.append(f"listener_log_tail={format_log_field(text_log_tail)}")
    fields.extend(create_libreoffice_runtime_guidance_fields())
    raise TimeoutError("\n".join(fields))


def create_property(name: str, value: object) -> Any:
    module_beans = importlib.import_module("com.sun.star.beans")
    prop = module_beans.PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def copy_docx_to_stage(file_in_docx: Path, *, dir_stage: Path) -> Path:
    file_staged_docx = dir_stage / file_in_docx.name
    shutil.copy2(file_in_docx, file_staged_docx)
    return file_staged_docx


def find_source_lock_file(file_in_docx: Path) -> Path | None:
    file_lock = file_in_docx.parent / f".~lock.{file_in_docx.name}#"
    if file_lock.exists():
        return file_lock
    return None


def open_listener_log_handle(
    file_listener_log: Path | None,
) -> tuple[
    Path | None,
    BufferedWriter | None,
    BufferedWriter | None,
    BufferedWriter | None,
]:
    if file_listener_log is None:
        return None, None, None, None
    file_listener_log = file_listener_log.resolve()
    file_listener_log.parent.mkdir(parents=True, exist_ok=True)
    handle_listener_log = file_listener_log.open("ab")
    return (
        file_listener_log,
        handle_listener_log,
        handle_listener_log,
        handle_listener_log,
    )


def connect_desktop(uno_module: Any, port: int) -> UnoDesktop:
    context_local = uno_module.getComponentContext()
    resolver = context_local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        context_local,
    )
    context_remote = resolver.resolve(
        f"uno:socket,host={LISTENER_HOST},port={port};urp;StarOffice.ComponentContext"
    )
    return context_remote.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop",
        context_remote,
    )


def refresh_uno_document_fields(doc: UnoTextDocument) -> None:
    doc.refresh()
    doc.updateLinks()
    indexes = doc.getDocumentIndexes()
    for idx in range(indexes.getCount()):
        indexes.getByIndex(idx).update()
    doc.getTextFields().refresh()
    doc.refresh()


def load_uno_document_or_raise(
    *,
    uno_module: Any,
    desktop: UnoDesktop,
    file_in_docx_source: Path,
    file_in_docx_staged: Path,
    exe_libreoffice: Path,
    dir_user_profile: Path,
    process_listener: ListenerProcess,
    file_listener_log: Path | None,
    file_source_lock: Path | None,
) -> UnoTextDocument:
    file_url = uno_module.systemPathToFileUrl(str(file_in_docx_staged.resolve()))
    props_default = (
        create_property("Hidden", True),
        create_property("ReadOnly", False),
        create_property("UpdateDocMode", 1),
    )
    props_hidden_only = (create_property("Hidden", True),)

    doc_probe = load_document_with_retry(
        desktop=desktop,
        url="private:factory/swriter",
        properties=props_hidden_only,
    )
    probe_ok = doc_probe is not None
    if doc_probe is not None:
        close_document(doc_probe)

    doc = load_document_with_retry(
        desktop=desktop,
        url=file_url,
        properties=props_default,
    )
    load_default_ok = doc is not None
    if doc is not None:
        return doc

    doc = load_document_with_retry(
        desktop=desktop,
        url=file_url,
        properties=props_hidden_only,
    )
    load_hidden_only_ok = doc is not None
    if doc is not None:
        return doc

    raise RuntimeError(
        "\n".join(
            create_load_failure_fields(
                file_in_docx_source=file_in_docx_source,
                file_in_docx_staged=file_in_docx_staged,
                file_url=file_url,
                exe_libreoffice=exe_libreoffice,
                dir_user_profile=dir_user_profile,
                process_listener=process_listener,
                file_listener_log=file_listener_log,
                file_source_lock=file_source_lock,
                probe_ok=probe_ok,
                load_default_ok=load_default_ok,
                load_hidden_only_ok=load_hidden_only_ok,
            )
        )
    )


def load_document_with_retry(
    *,
    desktop: UnoDesktop,
    url: str,
    properties: tuple[Any, ...],
) -> UnoTextDocument | None:
    deadline = time.monotonic() + DOCUMENT_LOAD_TIMEOUT_SECONDS
    while True:
        doc = desktop.loadComponentFromURL(url, "_blank", 0, properties)
        if doc is not None:
            return doc
        if time.monotonic() >= deadline:
            return None
        time.sleep(DOCUMENT_LOAD_POLL_INTERVAL_SECONDS)


def create_load_failure_fields(
    *,
    file_in_docx_source: Path,
    file_in_docx_staged: Path,
    file_url: str,
    exe_libreoffice: Path,
    dir_user_profile: Path,
    process_listener: ListenerProcess,
    file_listener_log: Path | None,
    file_source_lock: Path | None,
    probe_ok: bool,
    load_default_ok: bool,
    load_hidden_only_ok: bool,
) -> list[str]:
    exit_code = process_listener.poll()
    text_log_tail = read_log_tail(file_listener_log)

    if not probe_ok:
        reason_code = "uno_writer_probe_failed"
    elif exit_code not in (None, 0):
        reason_code = "listener_exited"
    elif not load_default_ok and not load_hidden_only_ok:
        reason_code = "staged_docx_import_failed"
    else:
        reason_code = "unknown_load_failure"

    fields = [
        "error_code=libreoffice_uno_load_failed",
        f"reason_code={reason_code}",
        f"file_in_docx={file_in_docx_source.resolve()}",
        f"file_in_docx_staged={file_in_docx_staged.resolve()}",
        f"file_url={file_url}",
        f"exe_libreoffice={exe_libreoffice.resolve()}",
        f"dir_user_profile={dir_user_profile.resolve()}",
        f"listener_exit_code={exit_code}",
        f"listener_log={listener_log_label(file_listener_log)}",
        f"source_lock_file_present={file_source_lock is not None}",
        f"probe_swriter_factory={'ok' if probe_ok else 'failed'}",
        f"load_staged_default_props={'ok' if load_default_ok else 'failed'}",
        f"load_staged_hidden_only={'ok' if load_hidden_only_ok else 'failed'}",
        f"staged_docx_size_bytes={file_in_docx_staged.stat().st_size}",
    ]
    if file_source_lock is not None:
        fields.append(f"source_lock_file={file_source_lock.resolve()}")
    if text_log_tail:
        fields.append(f"listener_log_tail={format_log_field(text_log_tail)}")
    fields.extend(create_libreoffice_runtime_guidance_fields())
    return fields


def read_log_tail(file_log: Path | None, *, max_bytes: int = 4000) -> str:
    if file_log is None or not file_log.exists():
        return ""
    with file_log.open("rb") as handle_log:
        handle_log.seek(0, 2)
        size = handle_log.tell()
        handle_log.seek(max(size - max_bytes, 0))
        return handle_log.read().decode("utf-8", errors="replace").strip()


def format_log_field(value: object) -> str:
    return str(value).replace("\n", r"\n")


def listener_log_label(file_listener_log: Path | None) -> str:
    if file_listener_log is None:
        return "stderr"
    return str(file_listener_log)


def close_document(doc: UnoTextDocument | None) -> None:
    if doc is None:
        return
    if hasattr(doc, "close"):
        doc.close(True)
    else:
        doc.dispose()


def terminate_process(process: ListenerProcess) -> None:
    process.terminate()
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=5)
