import re


class PhoneNormalizationError(ValueError):
    pass


PHONE_DIGITS_RE = re.compile(r"\d+")
    """Raised when phone normalization fails."""


NON_DIGIT_RE = re.compile(r"\D+")


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

    cleaned = str(raw_phone).strip()
    if not cleaned:
        raise PhoneNormalizationError("Phone is required.")

    cleaned = cleaned.replace(" ", "").replace("-", "")
    if cleaned.startswith("+86"):
        cleaned = cleaned[3:]
    elif cleaned.startswith("86") and len(cleaned) > 11:
        cleaned = cleaned[2:]

    digits = NON_DIGIT_RE.sub("", cleaned)
    if len(digits) != 11 or not digits.startswith("1"):
        raise PhoneNormalizationError(f"Invalid phone number: {raw_phone}")

    return digits
