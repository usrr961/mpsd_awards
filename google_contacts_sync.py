"""Legacy compatibility helpers; directory data is now imported from files."""

from importer import active_graduation_years


def calculate_active_grad_suffixes() -> tuple[str, ...]:
    return tuple(str(year)[-2:] for year in active_graduation_years())
