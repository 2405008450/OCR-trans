# -*- coding: utf-8 -*-
from __future__ import annotations

import locale
import os
import platform
import re
import shutil
import subprocess
import tempfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Optional

try:
    import ezdxf
    from ezdxf import recover
    from ezdxf.entities.acad_table import read_acad_table_content
    from ezdxf.tools.text import fast_plain_mtext
except ImportError:  # pragma: no cover - 依赖缺失时由运行时能力信息负责提示
    ezdxf = None
    recover = None
    read_acad_table_content = None
    fast_plain_mtext = None


CAD_EXTENSIONS = {".dwg", ".dws", ".dwt", ".dxf"}
ODA_REQUIRED_EXTENSIONS = {".dwg", ".dws", ".dwt"}
DEFAULT_ODA_TIMEOUT_SECONDS = 300


class CadTextError(RuntimeError):
    """CAD 文字提取失败。"""


class CadConverterUnavailableError(CadTextError):
    """当前环境缺少 ODA File Converter。"""


class CadConversionError(CadTextError):
    """ODA 转换失败。"""


class CadParseError(CadTextError):
    """DXF 解析失败。"""


@dataclass(frozen=True)
class CadTextItem:
    source_type: str
    source_label: str
    text: str
    paragraph_count: int = 1
    line_count: int = 1


@dataclass(frozen=True)
class CadExtractedContent:
    items: list[CadTextItem]
    page_count: int = 0
    paragraph_count: int = 0
    line_count: int = 0
    warning: str = ""
    stat_method: str = ""


def find_oda_file_converter(configured_path: str | None = None) -> Optional[Path]:
    """按配置、版本化安装目录和 PATH 的顺序查找 ODA File Converter。"""
    explicit = str(configured_path or os.getenv("ODA_FILE_CONVERTER_PATH", "")).strip().strip('"')
    if explicit:
        explicit_path = Path(explicit).expanduser()
        if explicit_path.is_file():
            return explicit_path.resolve()
        resolved = shutil.which(explicit)
        if resolved:
            return Path(resolved).resolve()

    candidates: list[Path] = []
    if platform.system() == "Windows":
        oda_root = Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "ODA"
        candidates.append(oda_root / "ODAFileConverter 27.1.0" / "ODAFileConverter.exe")
        if oda_root.is_dir():
            versioned = list(oda_root.glob("ODAFileConverter *"))
            versioned.sort(key=_version_key_from_path, reverse=True)
            candidates.extend(path / "ODAFileConverter.exe" for path in versioned)
        candidates.append(oda_root / "ODAFileConverter" / "ODAFileConverter.exe")

    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()

    for command in ("ODAFileConverter", "ODAFileConverter.exe"):
        resolved = shutil.which(command)
        if resolved:
            return Path(resolved).resolve()
    return None


def get_cad_support_info(configured_path: str | None = None) -> dict[str, Any]:
    parser_available = ezdxf is not None
    oda_path = find_oda_file_converter(configured_path)
    oda_available = oda_path is not None
    supported_extensions: set[str] = set()
    if parser_available:
        supported_extensions.add(".dxf")
        if oda_available:
            supported_extensions.update(ODA_REQUIRED_EXTENSIONS)

    reasons: list[str] = []
    if not parser_available:
        reasons.append("未安装 ezdxf，无法解析 DXF")
    if not oda_available:
        reasons.append("未找到 ODA File Converter，DWG/DWS/DWT 暂不可统计")

    return {
        "direct_dxf": parser_available,
        "oda_available": oda_available,
        "provider": "ezdxf + ODA File Converter",
        "version": _oda_version_from_path(oda_path) if oda_path else "",
        "supported_extensions": sorted(supported_extensions),
        "unavailable_reason": "；".join(reasons),
    }


def extract_cad_text(
    source_path: str | Path,
    *,
    workspace_dir: str | Path,
    oda_path: str | None = None,
    timeout_seconds: int = DEFAULT_ODA_TIMEOUT_SECONDS,
) -> CadExtractedContent:
    source = Path(source_path).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"CAD 文件不存在: {source}")
    extension = source.suffix.lower()
    if extension not in CAD_EXTENSIONS:
        raise ValueError(f"不支持的 CAD 文件格式: {extension or '无扩展名'}")
    if ezdxf is None:
        raise CadConverterUnavailableError("未安装 ezdxf，无法解析 CAD 文件")

    workspace = Path(workspace_dir).resolve()
    workspace.mkdir(parents=True, exist_ok=True)
    load_warnings: list[str] = []
    used_oda = extension in ODA_REQUIRED_EXTENSIONS

    if extension == ".dxf":
        try:
            document, direct_warning = _read_dxf_with_recovery(source)
            if direct_warning:
                load_warnings.append(direct_warning)
        except Exception as direct_exc:
            converter = find_oda_file_converter(oda_path)
            if converter is None:
                raise CadParseError(f"DXF 解析失败: {direct_exc}") from direct_exc
            used_oda = True
            document, oda_warning = _load_via_oda(
                source,
                workspace=workspace,
                converter=converter,
                timeout_seconds=timeout_seconds,
            )
            load_warnings.append(f"原始 DXF 解析失败，已通过 ODA 规范化: {direct_exc}")
            if oda_warning:
                load_warnings.append(oda_warning)
    else:
        converter = find_oda_file_converter(oda_path)
        if converter is None:
            raise CadConverterUnavailableError(
                "未找到 ODA File Converter；请配置 ODA_FILE_CONVERTER_PATH"
            )
        document, oda_warning = _load_via_oda(
            source,
            workspace=workspace,
            converter=converter,
            timeout_seconds=timeout_seconds,
        )
        if oda_warning:
            load_warnings.append(oda_warning)

    items, warning_counts, paper_layout_count = _extract_document_items(document)
    warning_text = _format_warnings(load_warnings, warning_counts)
    return CadExtractedContent(
        items=items,
        page_count=paper_layout_count,
        paragraph_count=sum(item.paragraph_count for item in items),
        line_count=sum(item.line_count for item in items),
        warning=warning_text,
        stat_method=(
            "ODA File Converter转DXF+ezdxf实体解析+Word近似计数"
            if used_oda
            else "DXF实体解析+Word近似计数"
        ),
    )


def _load_via_oda(
    source: Path,
    *,
    workspace: Path,
    converter: Path,
    timeout_seconds: int,
) -> tuple[Any, str]:
    safe_timeout = max(1, int(timeout_seconds or DEFAULT_ODA_TIMEOUT_SECONDS))
    workspace.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="cad-oda-", dir=str(workspace)) as temp_root:
        temp_dir = Path(temp_root)
        input_dir = temp_dir / "input"
        output_dir = temp_dir / "output"
        input_dir.mkdir()
        output_dir.mkdir()

        staged_extension = ".dxf" if source.suffix.lower() == ".dxf" else ".dwg"
        staged_source = input_dir / f"source{staged_extension}"
        _link_or_copy(source, staged_source)

        command = [
            str(converter),
            str(input_dir),
            str(output_dir),
            "ACAD2018",
            "DXF",
            "0",
            "1",
            staged_source.name,
        ]
        env = os.environ.copy()
        command = _headless_command(command, env)
        startupinfo = None
        creationflags = 0
        if platform.system() == "Windows":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        try:
            result = subprocess.run(
                command,
                shell=False,
                capture_output=True,
                text=True,
                encoding=locale.getpreferredencoding(False),
                errors="replace",
                timeout=safe_timeout,
                check=False,
                env=env,
                startupinfo=startupinfo,
                creationflags=creationflags,
            )
        except subprocess.TimeoutExpired as exc:
            raise CadConversionError(f"ODA 转换超时（{safe_timeout} 秒）: {source.name}") from exc
        except OSError as exc:
            raise CadConversionError(f"无法启动 ODA File Converter: {exc}") from exc

        output_file = _find_output_dxf(output_dir, staged_source.stem)
        diagnostics = _process_diagnostics(result.stdout, result.stderr)
        if output_file is None:
            detail = f"，{diagnostics}" if diagnostics else ""
            raise CadConversionError(
                f"ODA 未生成 DXF 文件（退出码 {result.returncode}）{detail}"
            )

        warning = ""
        if result.returncode != 0:
            warning = f"ODA 返回退出码 {result.returncode}，但已生成可解析 DXF"
            if diagnostics:
                warning = f"{warning}（{diagnostics}）"
        document, parse_warning = _read_dxf_with_recovery(output_file)
        if parse_warning:
            warning = "；".join(part for part in (warning, parse_warning) if part)
        return document, warning


def _read_dxf_with_recovery(path: Path) -> tuple[Any, str]:
    if ezdxf is None:
        raise CadParseError("未安装 ezdxf")
    try:
        return ezdxf.readfile(path), ""
    except Exception as normal_exc:
        if recover is None:
            raise CadParseError(str(normal_exc)) from normal_exc
        try:
            document, auditor = recover.readfile(path)
        except Exception as recover_exc:
            raise CadParseError(
                f"DXF 常规解析失败: {normal_exc}；恢复解析失败: {recover_exc}"
            ) from recover_exc
        error_count = len(getattr(auditor, "errors", []) or [])
        return document, f"DXF 已通过恢复模式读取（审计错误 {error_count} 项）"


def _extract_document_items(document: Any) -> tuple[list[CadTextItem], Counter[str], int]:
    items: list[CadTextItem] = []
    warnings: Counter[str] = Counter()
    paper_layout_count = 0
    for layout in document.layouts:
        layout_name = str(getattr(layout, "name", "") or "未命名布局")
        if layout_name.casefold() != "model":
            paper_layout_count += 1
        _extract_entities(
            layout,
            items=items,
            warnings=warnings,
            layout_name=layout_name,
            block_stack=(),
        )
    return items, warnings, paper_layout_count


def _extract_entities(
    entities: Iterable[Any],
    *,
    items: list[CadTextItem],
    warnings: Counter[str],
    layout_name: str,
    block_stack: tuple[str, ...],
) -> None:
    for entity in entities:
        try:
            entity_type = str(entity.dxftype()).upper()
        except Exception:
            warnings["proxy"] += 1
            continue

        if entity_type == "INSERT":
            _extract_insert(
                entity,
                items=items,
                warnings=warnings,
                layout_name=layout_name,
                block_stack=block_stack,
            )
        elif entity_type in {"TEXT", "ATTRIB"}:
            if entity_type == "ATTRIB" and _entity_is_invisible(entity):
                continue
            _append_item(
                items,
                text=_plain_single_line_text(entity),
                source_type=f"cad_{entity_type.lower()}",
                source_label=_source_label(layout_name, block_stack, entity_type),
            )
        elif entity_type == "ATTDEF":
            if _attdef_is_constant(entity) and not _entity_is_invisible(entity):
                _append_item(
                    items,
                    text=_plain_single_line_text(entity),
                    source_type="cad_attdef",
                    source_label=_source_label(layout_name, block_stack, entity_type),
                )
        elif entity_type == "MTEXT":
            _append_item(
                items,
                text=_plain_mtext_entity(entity),
                source_type="cad_mtext",
                source_label=_source_label(layout_name, block_stack, entity_type),
            )
        elif entity_type in {"MLEADER", "MULTILEADER"}:
            _extract_multileader(
                entity,
                items=items,
                warnings=warnings,
                layout_name=layout_name,
                block_stack=block_stack,
            )
        elif entity_type == "ACAD_TABLE":
            _extract_acad_table(
                entity,
                items=items,
                warnings=warnings,
                layout_name=layout_name,
                block_stack=block_stack,
            )
        elif entity_type in {"DIMENSION", "ARC_DIMENSION", "LARGE_RADIAL_DIMENSION"}:
            _extract_dimension(
                entity,
                items=items,
                warnings=warnings,
                layout_name=layout_name,
                block_stack=block_stack,
            )
        elif entity_type == "TOLERANCE":
            content = str(getattr(getattr(entity, "dxf", None), "content", "") or "")
            _append_item(
                items,
                text=_plain_mtext_content(content),
                source_type="cad_tolerance",
                source_label=_source_label(layout_name, block_stack, entity_type),
            )
        elif entity_type in {"ACAD_PROXY_ENTITY", "PROXYENTITY"}:
            warnings["proxy"] += 1


def _extract_insert(
    insert: Any,
    *,
    items: list[CadTextItem],
    warnings: Counter[str],
    layout_name: str,
    block_stack: tuple[str, ...],
) -> None:
    try:
        instances = list(insert.multi_insert()) if int(getattr(insert, "mcount", 1) or 1) > 1 else [insert]
    except Exception:
        instances = [insert]

    for instance_index, instance in enumerate(instances, start=1):
        block_name = str(getattr(getattr(instance, "dxf", None), "name", "") or "未命名块")
        instance_stack = block_stack + (block_name,)
        for attrib in list(getattr(instance, "attribs", []) or []):
            if _entity_is_invisible(attrib):
                continue
            suffix = f"；阵列实例 {instance_index}" if len(instances) > 1 else ""
            _append_item(
                items,
                text=_plain_single_line_text(attrib),
                source_type="cad_attrib",
                source_label=f"{_source_label(layout_name, instance_stack, 'ATTRIB')}{suffix}",
            )

        try:
            block = instance.block()
        except Exception:
            block = None
        if block is None:
            warnings["missing_block"] += 1
            continue
        block_entity = getattr(block, "block", None)
        if bool(getattr(block_entity, "is_xref", False)) or bool(getattr(block_entity, "is_xref_overlay", False)):
            warnings["xref"] += 1
            continue
        if block_name.casefold() in {name.casefold() for name in block_stack}:
            warnings["cycle"] += 1
            continue
        if len(instance_stack) > 32:
            warnings["depth"] += 1
            continue
        _extract_entities(
            block,
            items=items,
            warnings=warnings,
            layout_name=layout_name,
            block_stack=instance_stack,
        )


def _extract_multileader(
    entity: Any,
    *,
    items: list[CadTextItem],
    warnings: Counter[str],
    layout_name: str,
    block_stack: tuple[str, ...],
) -> None:
    before_count = len(items)
    try:
        virtual_entities = list(entity.virtual_entities())
    except Exception:
        virtual_entities = []
    if virtual_entities:
        _extract_entities(
            virtual_entities,
            items=items,
            warnings=warnings,
            layout_name=layout_name,
            block_stack=block_stack + ("MULTILEADER",),
        )
    if len(items) > before_count:
        return

    context = getattr(entity, "context", None)
    mtext = getattr(context, "mtext", None)
    if mtext is not None:
        _append_item(
            items,
            text=_plain_mtext_content(str(getattr(mtext, "default_content", "") or "")),
            source_type="cad_multileader",
            source_label=_source_label(layout_name, block_stack, "MULTILEADER"),
        )
    for attrib in list(getattr(entity, "block_attribs", []) or []):
        _append_item(
            items,
            text=str(getattr(attrib, "text", "") or ""),
            source_type="cad_multileader_attrib",
            source_label=_source_label(layout_name, block_stack, "MULTILEADER ATTRIB"),
        )


def _extract_acad_table(
    entity: Any,
    *,
    items: list[CadTextItem],
    warnings: Counter[str],
    layout_name: str,
    block_stack: tuple[str, ...],
) -> None:
    if read_acad_table_content is None:
        warnings["table"] += 1
        return
    try:
        rows = read_acad_table_content(entity)
    except Exception:
        warnings["table"] += 1
        return
    for row_index, row in enumerate(rows, start=1):
        for column_index, value in enumerate(row, start=1):
            _append_item(
                items,
                text=str(value or ""),
                source_type="cad_table",
                source_label=(
                    f"{_source_label(layout_name, block_stack, 'ACAD_TABLE')}；"
                    f"单元格 {row_index},{column_index}"
                ),
            )


def _extract_dimension(
    entity: Any,
    *,
    items: list[CadTextItem],
    warnings: Counter[str],
    layout_name: str,
    block_stack: tuple[str, ...],
) -> None:
    try:
        geometry_block = entity.get_geometry_block()
    except Exception:
        geometry_block = None
    if geometry_block is None:
        override = str(getattr(getattr(entity, "dxf", None), "text", "") or "")
        if override not in {"", "<>", " "}:
            _append_item(
                items,
                text=_plain_mtext_content(override),
                source_type="cad_dimension",
                source_label=_source_label(layout_name, block_stack, "DIMENSION"),
            )
        else:
            warnings["dimension"] += 1
        return
    block_name = str(getattr(geometry_block, "name", "DIMENSION") or "DIMENSION")
    _extract_entities(
        geometry_block,
        items=items,
        warnings=warnings,
        layout_name=layout_name,
        block_stack=block_stack + (block_name,),
    )


def _append_item(
    items: list[CadTextItem],
    *,
    text: str,
    source_type: str,
    source_label: str,
) -> None:
    normalized = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    if not normalized.strip():
        return
    lines = normalized.splitlines() or [normalized]
    paragraph_count = max(1, sum(1 for line in lines if line.strip()))
    items.append(
        CadTextItem(
            source_type=source_type,
            source_label=source_label,
            text=normalized,
            paragraph_count=paragraph_count,
            line_count=max(1, len(lines)),
        )
    )


def _plain_single_line_text(entity: Any) -> str:
    plain_text = getattr(entity, "plain_text", None)
    if callable(plain_text):
        try:
            return str(plain_text() or "")
        except Exception:
            pass
    return str(getattr(getattr(entity, "dxf", None), "text", "") or "")


def _plain_mtext_entity(entity: Any) -> str:
    all_columns = getattr(entity, "all_columns_plain_text", None)
    if callable(all_columns):
        try:
            return str(all_columns() or "")
        except Exception:
            pass
    plain_text = getattr(entity, "plain_text", None)
    if callable(plain_text):
        try:
            return str(plain_text(fast=False) or "")
        except TypeError:
            return str(plain_text() or "")
        except Exception:
            pass
    return _plain_mtext_content(str(getattr(entity, "text", "") or ""))


def _plain_mtext_content(content: str) -> str:
    if fast_plain_mtext is None:
        return str(content or "").replace("\\P", "\n")
    try:
        return str(fast_plain_mtext(str(content or "")) or "")
    except Exception:
        return str(content or "").replace("\\P", "\n")


def _entity_is_invisible(entity: Any) -> bool:
    value = getattr(entity, "is_invisible", False)
    try:
        return bool(value() if callable(value) else value)
    except Exception:
        return False


def _attdef_is_constant(entity: Any) -> bool:
    value = getattr(entity, "is_const", None)
    if value is not None:
        try:
            return bool(value() if callable(value) else value)
        except Exception:
            pass
    try:
        return bool(int(getattr(entity.dxf, "flags", 0) or 0) & 2)
    except Exception:
        return False


def _source_label(layout_name: str, block_stack: tuple[str, ...], entity_type: str) -> str:
    parts = [f"布局 {layout_name}"]
    if block_stack:
        parts.append(f"块 {' > '.join(block_stack)}")
    parts.append(f"实体 {entity_type}")
    return "；".join(parts)


def _format_warnings(load_warnings: list[str], warnings: Counter[str]) -> str:
    messages = [message for message in load_warnings if message]
    labels = {
        "xref": "跳过外部参照",
        "cycle": "跳过循环块引用",
        "depth": "跳过超过 32 层的块引用",
        "missing_block": "跳过缺失的块定义",
        "proxy": "跳过无法解析的代理实体",
        "table": "跳过无法解析的 CAD 表格",
        "dimension": "跳过无可提取显示文字的标注",
    }
    for key, label in labels.items():
        count = int(warnings.get(key, 0))
        if count:
            messages.append(f"{label} {count} 项")
    return "；".join(messages)


def _headless_command(command: list[str], env: dict[str, str]) -> list[str]:
    if platform.system() != "Linux" or env.get("DISPLAY"):
        return command
    xvfb_run = shutil.which("xvfb-run")
    if xvfb_run:
        return [xvfb_run, "-a", *command]
    env.setdefault("QT_QPA_PLATFORM", "offscreen")
    return command


def _link_or_copy(source: Path, target: Path) -> None:
    try:
        os.link(source, target)
    except OSError:
        shutil.copy2(source, target)


def _find_output_dxf(output_dir: Path, expected_stem: str) -> Optional[Path]:
    expected_name = f"{expected_stem}.dxf".casefold()
    files = [path for path in output_dir.iterdir() if path.is_file() and path.suffix.casefold() == ".dxf"]
    for path in files:
        if path.name.casefold() == expected_name:
            return path
    return files[0] if len(files) == 1 else None


def _process_diagnostics(stdout: str, stderr: str, limit: int = 1000) -> str:
    combined = "；".join(part.strip() for part in (stdout or "", stderr or "") if part.strip())
    combined = re.sub(r"\s+", " ", combined).strip()
    return combined[:limit]


def _version_key_from_path(path: Path) -> tuple[int, ...]:
    match = re.search(r"(\d+(?:\.\d+)+)", path.name)
    if not match:
        return ()
    return tuple(int(part) for part in match.group(1).split("."))


def _oda_version_from_path(path: Path) -> str:
    if platform.system() == "Windows":
        try:
            import win32api

            info = win32api.GetFileVersionInfo(str(path), "\\")
            ms = int(info["FileVersionMS"])
            ls = int(info["FileVersionLS"])
            return ".".join(
                str(part)
                for part in (ms >> 16, ms & 0xFFFF, ls >> 16, ls & 0xFFFF)
            )
        except Exception:
            pass
    for value in (path.parent.name, path.name):
        match = re.search(r"(\d+(?:\.\d+)+)", value)
        if match:
            return match.group(1)
    return ""
