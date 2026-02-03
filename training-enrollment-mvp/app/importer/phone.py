import re


class PhoneNormalizationError(ValueError):
    pass


PHONE_DIGITS_RE = re.compile(r"\d+")


def normalize_phone(raw_phone: str) -> str:
    if raw_phone is None:
        raise PhoneNormalizationError("Phone is required.")
    cleaned = str(raw_phone).strip()
    if not cleaned:
        raise PhoneNormalizationError("Phone is required.")
    cleaned = cleaned.replace(" ", "").replace("-", "")
    cleaned = cleaned.replace("+86", "")
    digits = "".join(PHONE_DIGITS_RE.findall(cleaned))
    if len(digits) < 7:
        raise PhoneNormalizationError(f"Invalid phone number: {raw_phone}")
    return digits
