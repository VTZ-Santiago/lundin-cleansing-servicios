"""CMP01 — re-exports the COMPLEMENTARIA migration categorisation rule.

The apply() function lives in src.segmentacion.rules.group3.m01.
Config for this domain uses validity_end as the primary date and
pending_planned_qty as the saldo indicator (pending_delivery_value
is not populated for contratos).
"""
from src.segmentacion.rules.group3.m01 import apply  # noqa: F401

__all__ = ["apply"]
