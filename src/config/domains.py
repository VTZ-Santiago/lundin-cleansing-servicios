from dataclasses import dataclass


@dataclass(frozen=True)
class DomainConfig:
    cli_name: str
    package_name: str
    input_subdir: str
    display_name: str
    output_suffix: str
    report_suffix: str


_DOMAIN_ALIASES: dict[str, str] = {
    "contratos": "contratos",
    "contrato": "contratos",
    "ordenes-compra": "ordenes_compra",
    "ordenes_compra": "ordenes_compra",
    "ordenescompra": "ordenes_compra",
    "po": "ordenes_compra",
}


DOMAIN_CONFIGS: dict[str, DomainConfig] = {
    "contratos": DomainConfig(
        cli_name="contratos",
        package_name="contratos",
        input_subdir="contratos",
        display_name="Contratos",
        output_suffix="contratos",
        report_suffix="contratos",
    ),
    "ordenes_compra": DomainConfig(
        cli_name="ordenes-compra",
        package_name="ordenes_compra",
        input_subdir="ordenes-compra",
        display_name="Ordenes de Compra",
        output_suffix="ordenes_compra",
        report_suffix="PO",
    ),
}


def normalize_domain_name(domain: str) -> str:
    key = domain.strip().lower()
    normalized = _DOMAIN_ALIASES.get(key)
    if normalized is None:
        supported = ", ".join(sorted(config.cli_name for config in DOMAIN_CONFIGS.values()))
        raise ValueError(f"Unknown domain: {domain}. Supported: {supported}")
    return normalized


def get_domain_config(domain: str) -> DomainConfig:
    return DOMAIN_CONFIGS[normalize_domain_name(domain)]