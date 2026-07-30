from dataclasses import dataclass


@dataclass(slots=True)
class RegistryEntry:

    protocol: str
    index: str
    organization: str
    excel_row: int
    scanned: bool = False