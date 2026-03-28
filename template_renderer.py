import base64
import io
import json
import math
import os
import re
import sys
import traceback
from pathlib import Path


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


def emit(message_type, level="info", message="", **extra):
    payload = {"type": message_type, "level": level, "message": message}
    payload.update(extra)
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def load_job():
    if len(sys.argv) > 1 and sys.argv[1]:
        return json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))

    raw = sys.stdin.buffer.read().decode("utf-8")
    if not raw.strip():
        raise ValueError("未收到任务参数")
    return json.loads(raw)


try:
    import cv2
    import numpy as np
    from PIL import Image, ImageDraw, ImageFont
except Exception as exc:
    emit("log", "error", f"缺少渲染依赖: {exc}")
    sys.exit(1)


def sanitize_name(value, fallback="item"):
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', "_", str(value or ""))
    cleaned = re.sub(r"\s+", " ", cleaned).strip().rstrip(". ")
    return cleaned or fallback


def load_cv_rgba(image_path):
    data = np.fromfile(str(image_path), dtype=np.uint8)
    if data.size == 0:
        raise ValueError(f"无法读取图像: {image_path}")
    image = cv2.imdecode(data, cv2.IMREAD_UNCHANGED)
    if image is None:
        raise ValueError(f"无法解码图像: {image_path}")
    if image.ndim == 2:
        image = cv2.cvtColor(image, cv2.COLOR_GRAY2BGRA)
    elif image.shape[2] == 3:
        image = cv2.cvtColor(image, cv2.COLOR_BGR2BGRA)
    elif image.shape[2] == 4:
        image = image.copy()
    else:
        raise ValueError(f"不支持的图像通道数: {image_path}")
    return image


def save_cv_jpg(image, output_path):
    if image.shape[2] == 4:
        alpha = image[:, :, 3:4].astype(np.float32) / 255.0
        rgb = image[:, :, :3].astype(np.float32)
        white = np.full_like(rgb, 255.0)
        blended = rgb * alpha + white * (1.0 - alpha)
        image = blended.astype(np.uint8)
    success, encoded = cv2.imencode(".jpg", image[:, :, :3], [int(cv2.IMWRITE_JPEG_QUALITY), 95])
    if not success:
        raise ValueError(f"无法编码输出图像: {output_path}")
    encoded.tofile(str(output_path))


def save_cv_png(image, output_path):
    success, encoded = cv2.imencode(".png", image)
    if not success:
        raise ValueError(f"无法编码输出图像: {output_path}")
    encoded.tofile(str(output_path))


def save_cv_image(image, output_path):
    suffix = str(Path(output_path).suffix).lower()
    if suffix == ".png":
        save_cv_png(image, output_path)
    else:
        save_cv_jpg(image, output_path)


def parse_point(item):
    if isinstance(item, dict):
        return float(item.get("x", 0)), float(item.get("y", 0))
    if isinstance(item, (list, tuple)) and len(item) >= 2:
        return float(item[0]), float(item[1])
    raise ValueError("透视点数据格式错误")


def parse_points(config):
    candidate = None
    for key in ("points", "vertices", "corners", "quad"):
        if key in config:
            candidate = config[key]
            break

    if candidate is None:
        ordered_keys = ("topLeft", "topRight", "bottomRight", "bottomLeft")
        if all(key in config for key in ordered_keys):
            candidate = [config[key] for key in ordered_keys]

    if isinstance(candidate, dict):
        ordered_keys = ("topLeft", "topRight", "bottomRight", "bottomLeft")
        candidate = [candidate[key] for key in ordered_keys if key in candidate]

    if not isinstance(candidate, (list, tuple)) or len(candidate) != 4:
        raise ValueError("config.json 缺少四个透视顶点坐标")

    return np.float32([parse_point(item) for item in candidate])


def parse_placement(config):
    placement = config.get("placement") if isinstance(config, dict) else None
    if not isinstance(placement, dict):
        placement = {}
    return {
        "scale": max(0.1, float(placement.get("scale", 1.0))),
        "offsetX": float(placement.get("offsetX", 0.0)),
        "offsetY": float(placement.get("offsetY", 0.0)),
        "rotation": float(placement.get("rotation", 0.0))
    }


def parse_effects(config):
    effects = config.get("effects") if isinstance(config, dict) else None
    if not isinstance(effects, dict):
        effects = {}
    design_blend_mode = str(effects.get("designBlendMode", "multiply")).strip().lower()
    if design_blend_mode not in {"normal", "multiply"}:
        design_blend_mode = "multiply"
    return {
        "designBlendMode": design_blend_mode,
        "designOpacity": max(0.0, min(3.0, float(effects.get("designOpacity", 1.0)))),
        "designBrightness": max(0.2, min(3.0, float(effects.get("designBrightness", 1.0)))),
        "textureOpacity": max(0.0, min(3.0, float(effects.get("textureOpacity", 1.0)))),
        "highlightOpacity": max(0.0, min(3.0, float(effects.get("highlightOpacity", 1.0)))),
        "autoMatchOrientation": bool(effects.get("autoMatchOrientation", True))
    }


def transform_points(points, placement):
    scale = max(0.1, float((placement or {}).get("scale", 1.0)))
    offset_x = float((placement or {}).get("offsetX", 0.0))
    offset_y = float((placement or {}).get("offsetY", 0.0))
    rotation_deg = float((placement or {}).get("rotation", 0.0))

    center = points.mean(axis=0)
    shifted = points - center
    shifted *= scale

    if rotation_deg:
        radians = math.radians(rotation_deg)
        cos_v = math.cos(radians)
        sin_v = math.sin(radians)
        rotation = np.float32([
            [cos_v, -sin_v],
            [sin_v, cos_v]
        ])
        shifted = shifted @ rotation.T

    translated = shifted + center + np.float32([offset_x, offset_y])
    return translated.astype(np.float32)


def estimate_quad_size(points):
    top = np.linalg.norm(points[1] - points[0])
    right = np.linalg.norm(points[2] - points[1])
    bottom = np.linalg.norm(points[3] - points[2])
    left = np.linalg.norm(points[0] - points[3])
    return max(1.0, (top + bottom) / 2.0), max(1.0, (left + right) / 2.0)


def should_auto_rotate_design(design, points, effects):
    if not effects.get("autoMatchOrientation", True):
        return False
    design_height, design_width = design.shape[:2]
    if not design_width or not design_height:
        return False
    target_width, target_height = estimate_quad_size(points)
    target_aspect = target_width / max(1.0, target_height)
    normal_aspect = design_width / max(1.0, float(design_height))
    rotated_aspect = design_height / max(1.0, float(design_width))
    normal_error = abs(math.log(normal_aspect / max(target_aspect, 1e-6)))
    rotated_error = abs(math.log(rotated_aspect / max(target_aspect, 1e-6)))
    return rotated_error + 0.05 < normal_error


def apply_brightness(image, factor):
    if abs(factor - 1.0) < 1e-6:
        return image
    result = image.copy()
    result[:, :, :3] = np.clip(result[:, :, :3].astype(np.float32) * factor, 0, 255).astype(np.uint8)
    return result


def apply_layer_alpha(image, opacity):
    opacity = max(0.0, min(1.0, float(opacity)))
    if opacity >= 0.999:
        return image
    result = image.copy()
    result[:, :, 3] = np.clip(result[:, :, 3].astype(np.float32) * opacity, 0, 255).astype(np.uint8)
    return result


def apply_blend_strength(bottom, top, strength, blend_fn):
    strength = max(0.0, float(strength))
    if strength <= 0:
        return bottom.copy()

    full_passes = int(math.floor(strength))
    fractional = strength - full_passes
    result = bottom.copy()

    for _ in range(full_passes):
        result = blend_fn(result, top)

    if fractional > 1e-6:
        result = blend_fn(result, apply_layer_alpha(top, fractional))

    return result


def apply_alpha_mask(image, mask):
    if mask.shape[2] == 4:
        mask_alpha = mask[:, :, 3]
        # When the mask is exported from PS as a white shape on transparent
        # background, transparent pixels often keep white RGB values. In that
        # case the alpha channel is the real clipping boundary.
        if np.any(mask_alpha < 255):
            effective_mask = mask_alpha
        else:
            effective_mask = cv2.cvtColor(mask[:, :, :3], cv2.COLOR_BGR2GRAY)
    else:
        effective_mask = cv2.cvtColor(mask[:, :, :3], cv2.COLOR_BGR2GRAY)

    result = image.copy()
    result[:, :, 3] = ((result[:, :, 3].astype(np.float32) * (effective_mask.astype(np.float32) / 255.0))).astype(np.uint8)
    return result


def alpha_composite(bottom, top):
    base_rgb = bottom[:, :, :3].astype(np.float32)
    base_alpha = bottom[:, :, 3:4].astype(np.float32) / 255.0
    top_rgb = top[:, :, :3].astype(np.float32)
    top_alpha = top[:, :, 3:4].astype(np.float32) / 255.0
    out_alpha = top_alpha + base_alpha * (1.0 - top_alpha)
    safe_alpha = np.clip(out_alpha, 1e-6, 1.0)
    out_rgb = (top_rgb * top_alpha + base_rgb * base_alpha * (1.0 - top_alpha)) / safe_alpha
    merged = np.dstack((np.clip(out_rgb, 0, 255), np.clip(out_alpha * 255.0, 0, 255))).astype(np.uint8)
    return merged


def multiply_blend(bottom, top):
    base_rgb = bottom[:, :, :3].astype(np.float32)
    top_rgb = top[:, :, :3].astype(np.float32)
    if top.shape[2] == 4:
        top_alpha = top[:, :, 3:4].astype(np.float32) / 255.0
    else:
        top_alpha = np.ones((top.shape[0], top.shape[1], 1), dtype=np.float32)

    multiplied = (base_rgb * top_rgb) / 255.0
    mixed_rgb = base_rgb * (1.0 - top_alpha) + multiplied * top_alpha

    result = bottom.copy()
    result[:, :, :3] = np.clip(mixed_rgb, 0, 255).astype(np.uint8)
    return result


def normal_blend(bottom, top):
    return alpha_composite(bottom, top)


def screen_blend(bottom, top):
    base_rgb = bottom[:, :, :3].astype(np.float32)
    top_rgb = top[:, :, :3].astype(np.float32)
    if top.shape[2] == 4:
        top_alpha = top[:, :, 3:4].astype(np.float32) / 255.0
    else:
        top_alpha = np.ones((top.shape[0], top.shape[1], 1), dtype=np.float32)

    screened = 255.0 - ((255.0 - base_rgb) * (255.0 - top_rgb) / 255.0)
    mixed_rgb = base_rgb * (1.0 - top_alpha) + screened * top_alpha

    result = bottom.copy()
    result[:, :, :3] = np.clip(mixed_rgb, 0, 255).astype(np.uint8)
    return result


def decode_data_url(data_url):
    if not data_url or "," not in data_url:
        raise ValueError("水印图片数据无效")
    _, encoded = data_url.split(",", 1)
    return base64.b64decode(encoded)


def load_preset_image(preset):
    image_path = preset.get("imagePath")
    if image_path and os.path.exists(image_path):
        return Image.open(image_path).convert("RGBA")

    image_data_url = preset.get("imageDataUrl")
    if image_data_url:
        return Image.open(io.BytesIO(decode_data_url(image_data_url))).convert("RGBA")

    raise ValueError("图片水印缺少可用图像")


def load_font(font_size):
    candidates = [
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/arial.ttf"
    ]
    for font_path in candidates:
        if os.path.exists(font_path):
            try:
                return ImageFont.truetype(font_path, font_size)
            except Exception:
                pass
    return ImageFont.load_default()


def build_watermark_sprite(preset):
    scale = max(0.1, float(preset.get("scale", 1)))
    opacity = max(0.0, min(1.0, float(preset.get("opacity", 1))))
    rotation = float(preset.get("rotation", 0))
    content_type = preset.get("contentType") or preset.get("type") or "text"

    if content_type == "image":
        sprite = load_preset_image(preset)
    else:
        text = str(preset.get("text") or preset.get("content") or "Sample")
        font_size = max(12, int(preset.get("fontSize", 44)))
        font = load_font(font_size)
        probe = Image.new("RGBA", (8, 8), (0, 0, 0, 0))
        draw = ImageDraw.Draw(probe)
        bbox = draw.textbbox((0, 0), text, font=font)
        width = max(1, bbox[2] - bbox[0] + 24)
        height = max(1, bbox[3] - bbox[1] + 24)
        sprite = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(sprite)
        color = tuple(preset.get("color", [255, 255, 255]))[:3]
        draw.text((12 - bbox[0], 12 - bbox[1]), text, font=font, fill=(*color, 255))

    scaled_size = (
        max(1, int(sprite.width * scale)),
        max(1, int(sprite.height * scale))
    )
    sprite = sprite.resize(scaled_size, Image.Resampling.LANCZOS)

    if rotation:
        sprite = sprite.rotate(-rotation, expand=True, resample=Image.Resampling.BICUBIC)

    if opacity < 1:
        alpha = sprite.getchannel("A").point(lambda value: int(value * opacity))
        sprite.putalpha(alpha)

    return sprite


def build_watermark_layer(size, preset):
    if not preset:
        return None

    width, height = size
    layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    sprite = build_watermark_sprite(preset)
    mode = preset.get("layoutMode") or preset.get("mode") or "single"
    pos_x = float(preset.get("xRatio", 0.5)) * width
    pos_y = float(preset.get("yRatio", 0.5)) * height

    def paste_sprite(target_x, target_y):
        left = int(target_x - sprite.width / 2)
        top = int(target_y - sprite.height / 2)
        layer.alpha_composite(sprite, (left, top))

    if mode == "matrix":
        spacing_x = float(preset.get("spacingX", sprite.width * 1.4))
        spacing_y = float(preset.get("spacingY", sprite.height * 1.6))
        offset_x = float(preset.get("offsetX", 0))
        offset_y = float(preset.get("offsetY", 0))
        stagger = bool(preset.get("stagger", False))
        stagger_offset = float(preset.get("staggerOffset", spacing_x / 2 if spacing_x else sprite.width / 2))
        start_y = -sprite.height + offset_y
        row_index = 0
        while start_y < height + sprite.height:
            row_x = -sprite.width + offset_x
            if stagger and row_index % 2 == 1:
                row_x += stagger_offset
            while row_x < width + sprite.width:
                paste_sprite(row_x, start_y)
                row_x += max(1.0, spacing_x)
            start_y += max(1.0, spacing_y)
            row_index += 1
    else:
        paste_sprite(pos_x, pos_y)

    return layer


def pil_to_cv(image):
    return cv2.cvtColor(np.array(image), cv2.COLOR_RGBA2BGRA)


def get_scene_label(scene):
    scene_name = str(scene.get("name") or scene.get("label") or "").strip()
    if scene_name:
        return scene_name
    relative_path = str(scene.get("relativePath") or "").strip().replace("\\", "/").strip("/")
    if relative_path:
        return Path(relative_path).name
    return "scene"


def render_design_to_template(design_path, template_dir, preset, output_dir, *, config_override=None, output_name=None):
    design_path = Path(design_path)
    template_dir = Path(template_dir)

    base = load_cv_rgba(template_dir / "base.png")
    mask = load_cv_rgba(template_dir / "mask.png")
    config = json.loads((template_dir / "config.json").read_text(encoding="utf-8"))
    if isinstance(config_override, dict):
        config = {
            **config,
            **config_override
        }
    texture_path = template_dir / "texture.png"
    highlight_path = template_dir / "highlight.png"

    points = transform_points(parse_points(config), parse_placement(config))
    effects = parse_effects(config)
    design = load_cv_rgba(design_path)
    if should_auto_rotate_design(design, points, effects):
        design = cv2.rotate(design, cv2.ROTATE_90_CLOCKWISE)
    height, width = base.shape[:2]

    source = np.float32([
        [0, 0],
        [design.shape[1] - 1, 0],
        [design.shape[1] - 1, design.shape[0] - 1],
        [0, design.shape[0] - 1]
    ])
    transform = cv2.getPerspectiveTransform(source, points)
    warped = cv2.warpPerspective(
        design,
        transform,
        (width, height),
        flags=cv2.INTER_LINEAR,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=(0, 0, 0, 0)
    )

    warped = apply_brightness(warped, effects["designBrightness"])
    masked = apply_alpha_mask(warped, mask)
    design_blend_fn = multiply_blend if effects["designBlendMode"] == "multiply" else normal_blend
    composed = apply_blend_strength(base, masked, effects["designOpacity"], design_blend_fn)

    if texture_path.exists():
        texture = load_cv_rgba(texture_path)
        composed = apply_blend_strength(composed, texture, effects["textureOpacity"], multiply_blend)

    if highlight_path.exists():
        highlight = load_cv_rgba(highlight_path)
        composed = apply_blend_strength(composed, highlight, effects["highlightOpacity"], screen_blend)

    watermark_layer = build_watermark_layer((width, height), preset)
    if watermark_layer:
        composed = alpha_composite(composed, pil_to_cv(watermark_layer))

    output_path = Path(output_dir) / (
        output_name
        or f"{sanitize_name(design_path.stem, 'design')}_{sanitize_name(template_dir.name, 'template')}.jpg"
    )
    save_cv_image(composed, output_path)
    return str(output_path)


def main():
    job = load_job()
    template_root = Path(job.get("templateRootDir") or ".")
    output_dir = Path(job.get("outputDir") or ".")
    output_dir.mkdir(parents=True, exist_ok=True)

    designs = job.get("designs") or []
    preset = job.get("watermarkPreset")
    effect_preset = job.get("effectPreset")
    mode = str(job.get("mode") or "batch").strip().lower()
    processed = 0
    failed = 0

    if not designs:
        raise ValueError("没有需要处理的设计图")

    if mode == "preview":
        preview = job.get("preview") or {}
        design = designs[0] if designs else {}
        design_path = design.get("path") or ""
        if not design_path or not os.path.exists(design_path):
            raise ValueError("预览设计图不存在")

        relative_path = str(preview.get("relativePath") or "").replace("\\", "/").strip("/")
        if not relative_path:
            raise ValueError("缺少预览场景路径")

        template_dir = template_root / relative_path
        if not template_dir.exists():
            raise ValueError(f"预览场景不存在: {relative_path}")

        scene_name = get_scene_label(preview)
        design_name = sanitize_name(Path(design_path).stem, "design")
        preview_key = sanitize_name(str(preview.get("previewKey") or "").strip(), "")
        key_suffix = f"_{preview_key}" if preview_key else ""
        output_name = f"{design_name}_{sanitize_name(scene_name, 'preview')}{key_suffix}_preview.png"
        output_path = render_design_to_template(
            design_path,
            template_dir,
            preset,
            output_dir,
            config_override={
                "placement": preview.get("placement") or {},
                "effects": preview.get("effects") or effect_preset or {}
            },
            output_name=output_name
        )
        emit("done", "success", "预览渲染完成", outputPath=output_path, preview=True)
        return

    template_groups = job.get("templateGroups") or []
    if not template_groups:
        legacy_templates = job.get("templates") or []
        template_groups = [
            {
                "name": template_name,
                "scenes": [
                    {
                        "name": template_name,
                        "relativePath": template_name
                    }
                ]
            }
            for template_name in legacy_templates
        ]

    if not template_groups:
        raise ValueError("没有选择模板组")

    total = len(designs) * sum(
        max(1, len(group.get("scenes") or []))
        for group in template_groups
    )

    for group in template_groups:
        group_name = str(group.get("name") or "未命名模板组").strip() or "未命名模板组"
        scenes = group.get("scenes") or []
        if not scenes:
            failed += len(designs)
            emit("log", "error", f"模板组 {group_name} 未包含可用场景")
            continue

        emit("log", "info", f"开始处理模板组: {group_name}")
        for design in designs:
            design_name = design.get("name") or "未命名设计图"
            design_path = design.get("path") or ""
            if not design_path or not os.path.exists(design_path):
                failed += len(scenes)
                emit("log", "warn", f"跳过 {design_name}: 未找到本地文件路径")
                continue

            design_folder_name = sanitize_name(Path(design_path).stem, "design")
            group_folder_name = sanitize_name(group_name, "template-group")
            design_output_dir = output_dir / f"{design_folder_name}_{group_folder_name}"
            design_output_dir.mkdir(parents=True, exist_ok=True)

            for scene in scenes:
                scene_relative_path = str(scene.get("relativePath") or "").replace("\\", "/").strip("/")
                scene_name = get_scene_label(scene)
                if not scene_relative_path:
                    failed += 1
                    emit("log", "error", f"{design_name} -> {group_name}/{scene_name} 失败: 缺少场景路径")
                    continue

                template_dir = template_root / scene_relative_path
                required = ["base.png", "mask.png", "config.json"]
                missing = [name for name in required if not (template_dir / name).exists()]
                has_texture_stack = (template_dir / "texture.png").exists() and (template_dir / "highlight.png").exists()
                if not has_texture_stack:
                    missing.extend(["texture.png + highlight.png"])
                if missing:
                    failed += 1
                    emit("log", "error", f"模板 {group_name}/{scene_name} 缺少文件: {', '.join(missing)}")
                    continue

                try:
                    output_path = render_design_to_template(
                        design_path,
                        template_dir,
                        preset,
                        design_output_dir,
                        config_override={"effects": effect_preset or {}},
                        output_name=f"{design_folder_name}_{sanitize_name(scene_name, 'scene')}.jpg"
                    )
                    processed += 1
                    emit(
                        "log",
                        "success",
                        f"{design_name} -> {group_name}/{scene_name} 已生成",
                        outputPath=output_path,
                        outputDir=str(design_output_dir),
                        groupName=group_name,
                        sceneName=scene_name,
                        designName=design_name,
                        designPath=design_path
                    )
                    emit("progress", "info", "进度更新", processed=processed, failed=failed, total=total)
                except Exception as exc:
                    failed += 1
                    emit("log", "error", f"{design_name} -> {group_name}/{scene_name} 失败: {exc}")

    emit("done", "success", "智能模板任务完成", processed=processed, failed=failed, total=total)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        emit("log", "warn", "智能模板任务已取消")
        sys.exit(1)
    except Exception as exc:
        emit("log", "error", f"智能模板任务异常: {exc}")
        emit("log", "error", traceback.format_exc())
        sys.exit(1)
