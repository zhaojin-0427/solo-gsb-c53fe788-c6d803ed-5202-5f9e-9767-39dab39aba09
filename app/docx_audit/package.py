"""DOCX package loading: unzip, encryption checks and relationship lookup."""

from __future__ import annotations

import io
import zipfile

from lxml import etree

from .errors import AuditError
from .namespaces import PKG_REL_NS

ENCRYPTED_CONTENT_TYPES = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.encrypted",
)


def load_package(data: bytes) -> "Package":
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise AuditError(
            "not_a_docx", "上传内容不是有效的 DOCX（OOXML ZIP）文件"
        )
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise AuditError(
            "not_a_docx", "上传内容不是有效的 DOCX（OOXML ZIP）文件"
        )

    names = set(zf.namelist())
    if "[Content_Types].xml" not in names:
        raise AuditError("not_a_docx", "DOCX 缺少 [Content_Types].xml")

    try:
        content_types = etree.fromstring(zf.read("[Content_Types].xml"))
    except (KeyError, etree.XMLSyntaxError):
        raise AuditError("corrupt_docx", "无法解析 [Content_Types].xml")

    for override in content_types.iter():
        ct = override.get("ContentType", "")
        if ct in ENCRYPTED_CONTENT_TYPES:
            raise AuditError("encrypted_docx", "不支持加密 DOCX，请先解除密码")
    if names & {"EncryptedPackage", "EncryptionInfo"}:
        raise AuditError("encrypted_docx", "不支持加密 DOCX，请先解除密码")

    return Package(zf)


class Package:
    def __init__(self, zf: zipfile.ZipFile):
        self._zf = zf
        self._main_part = self._resolve_main_part()

    def _resolve_main_part(self):
        rels_path = "_rels/.rels"
        try:
            rels = etree.fromstring(self._zf.read(rels_path))
        except (KeyError, etree.XMLSyntaxError):
            if "word/document.xml" in self._zf.namelist():
                return "word/document.xml"
            raise AuditError("corrupt_docx", "DOCX 缺少包关系与 word/document.xml")

        for rel in rels.iter(f"{{{PKG_REL_NS}}}Relationship"):
            if rel.get("Type", "").endswith("/officeDocument"):
                target = rel.get("Target", "")
                return target.lstrip("/")
        if "word/document.xml" in self._zf.namelist():
            return "word/document.xml"
        raise AuditError("corrupt_docx", "DOCX 关系中找不到主文档部件")

    def _part_xml(self, path):
        try:
            return etree.fromstring(self._zf.read(path))
        except KeyError:
            return None
        except etree.XMLSyntaxError:
            raise AuditError("corrupt_docx", f"无法解析 {path}")

    def document(self):
        root = self._part_xml(self._main_part)
        if root is None:
            raise AuditError("corrupt_docx", f"缺少主文档部件 {self._main_part}")
        return root

    def part_dir(self):
        return self._main_part.rsplit("/", 1)[0] if "/" in self._main_part else ""

    def resolve_relative_part(self, target):
        """Resolve a relationship target relative to the main document part."""
        if target.startswith("/"):
            return target.lstrip("/")
        base = self.part_dir()
        if base:
            return f"{base}/{target}"
        return target

    def numbering(self):
        return self._related_part("numbering")

    def styles(self):
        return self._related_part("styles")

    def _related_part(self, suffix):
        doc_path = self._main_part
        base_dir = self.part_dir()
        rels_path = f"{base_dir}/_rels/{doc_path.rsplit('/', 1)[-1]}.rels"
        if rels_path in self._zf.namelist():
            try:
                rels = etree.fromstring(self._zf.read(rels_path))
            except etree.XMLSyntaxError:
                raise AuditError("corrupt_docx", f"无法解析 {rels_path}")
            for rel in rels.iter(f"{{{PKG_REL_NS}}}Relationship"):
                if rel.get("Type", "").endswith(f"/{suffix}"):
                    return self._part_xml(
                        self.resolve_relative_part(rel.get("Target", ""))
                    )
        # Default locations used by Word.
        default = f"{base_dir}/{suffix}.xml" if base_dir else f"{suffix}.xml"
        if default in self._zf.namelist():
            return self._part_xml(default)
        return None
