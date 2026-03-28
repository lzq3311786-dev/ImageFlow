import json
import sys
from pathlib import Path


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


try:
    import cv2
    import numpy as np
except Exception:
    sys.exit(1)


def order_points(points):
    pts = np.array(points, dtype=np.float32)
    sums = pts.sum(axis=1)
    diffs = pts[:, 0] - pts[:, 1]
    ordered = np.zeros((4, 2), dtype=np.float32)
    ordered[0] = pts[np.argmin(sums)]
    ordered[2] = pts[np.argmax(sums)]
    ordered[1] = pts[np.argmax(diffs)]
    ordered[3] = pts[np.argmin(diffs)]
    return ordered


def scale_points(points, scale=1.05):
    center = points.mean(axis=0)
    return (center + (points - center) * scale).round().astype(np.int32)


def main():
    raw = sys.stdin.buffer.read().decode("utf-8")
    payload = json.loads(raw or "{}")
    mask_path = Path(str(payload.get("maskPath") or "").strip())
    if not mask_path.exists():
        sys.exit(1)

    image = cv2.imdecode(np.fromfile(str(mask_path), dtype=np.uint8), cv2.IMREAD_UNCHANGED)
    if image is None:
        sys.exit(1)

    if image.ndim == 2:
        gray = image
        alpha = None
    elif image.shape[2] == 4:
        alpha = image[:, :, 3]
        gray = cv2.cvtColor(image[:, :, :3], cv2.COLOR_BGR2GRAY)
    else:
        alpha = None
        gray = cv2.cvtColor(image[:, :, :3], cv2.COLOR_BGR2GRAY)

    if alpha is not None:
        _, alpha_mask = cv2.threshold(alpha, 8, 255, cv2.THRESH_BINARY)
        _, gray_mask = cv2.threshold(gray, 12, 255, cv2.THRESH_BINARY)
        binary = cv2.bitwise_and(alpha_mask, gray_mask)
    else:
        _, binary = cv2.threshold(gray, 12, 255, cv2.THRESH_BINARY)

    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = [cnt for cnt in contours if cv2.contourArea(cnt) > 16]
    if not contours:
        sys.exit(1)

    contour = max(contours, key=cv2.contourArea)
    perimeter = cv2.arcLength(contour, True)
    quad = None
    for epsilon_ratio in (0.01, 0.015, 0.02, 0.03, 0.04):
        approx = cv2.approxPolyDP(contour, perimeter * epsilon_ratio, True)
        if len(approx) == 4:
            quad = approx.reshape(4, 2)
            break

    if quad is None:
        rect = cv2.minAreaRect(contour)
        quad = cv2.boxPoints(rect)

    ordered = order_points(quad)
    scaled = scale_points(ordered, 1.05)
    result = {
        "points": [
            {"x": int(point[0]), "y": int(point[1])}
            for point in scaled
        ]
    }
    sys.stdout.write(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
