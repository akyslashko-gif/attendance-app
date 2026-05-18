"""Bulgarian public holidays with weekend substitution rule."""
from datetime import date, timedelta


def orthodox_easter(year):
    """Return Orthodox Easter Sunday (Julian calendar projected to Gregorian)."""
    a = year % 4
    b = year % 7
    c = year % 19
    d = (19 * c + 15) % 30
    e = (2 * a + 4 * b - d + 34) % 7
    month = (d + e + 114) // 31
    day   = ((d + e + 114) % 31) + 1
    return date(year, month, day) + timedelta(days=13)  # Julian → Gregorian


def get_bg_holidays(year):
    """
    Return dict of {date_iso: (name, is_substitute)} for Bulgarian public holidays.
    Applies the rule: if a holiday falls on a weekend, the next available
    working day becomes a substitute day off.

    Easter cycle is excluded from substitution — it already includes Monday.
    """
    easter = orthodox_easter(year)

    fixed = [
        date(year,  1,  1),  # Нова година
        date(year,  3,  3),  # Ден на Освобождението
        date(year,  5,  1),  # Ден на труда
        date(year,  5,  6),  # Гергьовден
        date(year,  5, 24),  # Ден на просветата
        date(year,  9,  6),  # Ден на Съединението
        date(year,  9, 22),  # Ден на Независимостта
        date(year, 11,  1),  # Ден на народните будители
        date(year, 12, 24),  # Бъдни вечер
        date(year, 12, 25),  # Рождество Христово
        date(year, 12, 26),  # Рождество Христово
    ]

    names = {
        (1,  1):  "Нова година",
        (3,  3):  "Ден на Освобождението",
        (5,  1):  "Ден на труда",
        (5,  6):  "Гергьовден",
        (5, 24):  "Ден на просветата и културата",
        (9,  6):  "Ден на Съединението",
        (9, 22):  "Ден на Независимостта",
        (11, 1):  "Ден на народните будители",
        (12, 24): "Бъдни вечер",
        (12, 25): "Рождество Христово",
        (12, 26): "Рождество Христово",
    }

    # Start with all holidays (fixed + Easter cycle)
    holidays = {}  # iso → (display_name, is_substitute)

    for d in fixed:
        holidays[d.isoformat()] = (names[(d.month, d.day)], False)

    # Easter cycle — never substituted (Monday already included)
    holidays[(easter - timedelta(days=2)).isoformat()] = ("Разпети петък", False)
    holidays[(easter - timedelta(days=1)).isoformat()] = ("Велика събота", False)
    holidays[easter.isoformat()]                        = ("Великден", False)
    holidays[(easter + timedelta(days=1)).isoformat()]  = ("Великденски понеделник", False)

    # Weekend substitution: fixed holidays only (not Easter cycle)
    substitutes = {}
    for d in fixed:
        if d.weekday() >= 5:  # Saturday=5, Sunday=6
            # Find the next working day not already occupied
            substitute = d + timedelta(days=(7 - d.weekday()))  # next Monday
            while (substitute.isoformat() in holidays or
                   substitute.isoformat() in substitutes):
                substitute += timedelta(days=1)
            original_name = names[(d.month, d.day)]
            substitutes[substitute.isoformat()] = (
                f"{original_name} (почивен ден)", True
            )

    holidays.update(substitutes)
    return holidays


def get_bg_holidays_range(year_from, year_to):
    """Merged holiday dict for a range of years."""
    result = {}
    for y in range(year_from, year_to + 1):
        result.update(get_bg_holidays(y))
    return result


def holiday_name(holidays, iso):
    """Return display name for a holiday date, or empty string."""
    entry = holidays.get(iso)
    return entry[0] if entry else ""


def is_substitute(holidays, iso):
    """Return True if this date is a substitute working day."""
    entry = holidays.get(iso)
    return entry[1] if entry else False
