from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Optional


class SelectorType(str, Enum):
    EMAIL    = "email"
    PHONE    = "phone"
    IP       = "ip"
    ADDRESS  = "address"
    LINKEDIN = "linkedin"
    GITHUB   = "github"
    TELEGRAM = "telegram"
    DISCORD  = "discord"
    NAME     = "name"
    OTHER    = "other"


@dataclass
class Selector:
    selector_id: str
    selector: str
    selector_clean: str
    selector_type: SelectorType
    target_id: Optional[str]
    nork_id: Optional[str]
    date_created: str
    created_by: str
    last_updated: str
    last_updated_by: str
    data_source: Optional[str]


@dataclass
class Target:
    target_id: str
    target_name: Optional[str]
    case_number: Optional[str]
    laptop_count: int
    date_created: str
    created_by: str
    last_updated: str
    last_updated_by: str
    data_source: Optional[str]


@dataclass(frozen=True)
class ColumnTypeInfo:
    col_index: int
    selector_type: SelectorType
    confidence: float


@dataclass
class GWResult:
    query_value: str
    selector: str
    selector_id: Optional[str]
    selector_clean: Optional[str]
    selector_type: Optional[SelectorType]
    target_id: Optional[str]
    nork_id: Optional[str]
    date_created: Optional[str]
    created_by: Optional[str]
    last_updated: Optional[str]
    last_updated_by: Optional[str]
    data_source: Optional[str]
    target_name: Optional[str]
    in_gray_wolfe: bool


@dataclass
class SApiResult:
    s_id: str
    selector: str
    doc_id: str
    doc_type: str
    doc_sub_type: str
    case: str
    serial: str
    case_serial_full: str
    office: str
    doc_title: str
    author: str
    created_date: str
    link: str
