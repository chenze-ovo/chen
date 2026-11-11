# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path


def package_root() -> Path:
    return Path(__file__).resolve().parent


def builtin_templates_dir() -> Path:
    return package_root() / "templates"


def builtin_rules_path() -> Path:
    return package_root() / "resources" / "rules.yaml"


def default_template_d01() -> Path:
    return builtin_templates_dir() / "dian_01_blank.xls"


def default_template_d02() -> Path:
    return builtin_templates_dir() / "dian_02_blank.xls"


def default_template_dx() -> Path:
    return builtin_templates_dir() / "dianxiaomi_blank.xlsx"
