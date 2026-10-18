"""Employee photo: validate an upload and normalize it to a small, metadata-free JPEG.

The photo is employee data: it lives only in memory with the report (never on disk, never logged).
It is never stretched — every output fits it inside its box preserving the aspect ratio.
"""
from __future__ import annotations

import io
from dataclasses import dataclass

from PIL import Image, ImageOps, UnidentifiedImageError

MAX_UPLOAD_BYTES = 15 * 1024 * 1024
MAX_SIDE = 1200            # longest side after normalization (plenty for print at photo-box size)
MIN_SIDE = 80
ACCEPTED_FORMATS = {"JPEG", "PNG", "WEBP", "BMP", "GIF"}
Image.MAX_IMAGE_PIXELS = 60_000_000   # refuse decompression bombs


class PhotoError(Exception):
    """Persian, user-facing."""


@dataclass(frozen=True)
class Photo:
    jpeg: bytes
    width: int
    height: int


def process_upload(data: bytes) -> Photo:
    if not data:
        raise PhotoError("فایلی انتخاب نشده است.")
    if len(data) > MAX_UPLOAD_BYTES:
        raise PhotoError("حجم عکس بیش از ۱۵ مگابایت است. عکس کوچک‌تری انتخاب کنید.")
    try:
        img = Image.open(io.BytesIO(data))
        fmt = img.format
        img.load()
    except (UnidentifiedImageError, OSError, Image.DecompressionBombError, ValueError):
        raise PhotoError("فایل انتخاب‌شده عکس معتبری نیست. یک عکس JPG یا PNG انتخاب کنید.")
    if fmt not in ACCEPTED_FORMATS:
        raise PhotoError("این قالب عکس پشتیبانی نمی‌شود. یک عکس JPG یا PNG انتخاب کنید.")
    if getattr(img, "n_frames", 1) > 1:
        img.seek(0)

    img = ImageOps.exif_transpose(img)          # phone photos: apply the camera's rotation
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, "white")  # transparent areas become white, not black
        bg.paste(img, mask=img.getchannel("A"))
        img = bg
    else:
        img = img.convert("RGB")

    if min(img.size) < MIN_SIDE:
        raise PhotoError("ابعاد عکس خیلی کوچک است. عکسی با کیفیت بهتر انتخاب کنید.")
    img.thumbnail((MAX_SIDE, MAX_SIDE), Image.Resampling.LANCZOS)   # keeps aspect ratio

    out = io.BytesIO()
    img.save(out, "JPEG", quality=88, optimize=True)   # re-encoding drops EXIF/GPS metadata
    return Photo(out.getvalue(), img.width, img.height)


def fit(box_w: float, box_h: float, img_w: int, img_h: int, pad: float = 0) -> tuple[float, float, float, float]:
    """Largest size that fits the box without distortion ("contain"), centred. Returns (x, y, w, h)."""
    avail_w, avail_h = box_w - 2 * pad, box_h - 2 * pad
    scale = min(avail_w / img_w, avail_h / img_h)
    w, h = img_w * scale, img_h * scale
    return (box_w - w) / 2, (box_h - h) / 2, w, h
