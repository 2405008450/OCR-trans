import sys
import zipfile
from pathlib import Path

from docx import Document
from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.service.web_asset_preserve_docx_service import render_web_asset_preserve_docx_from_layout


def test_render_web_asset_preserve_docx_keeps_text_and_important_images(tmp_path):
    image_path = tmp_path / "web.png"
    image = Image.new("RGB", (520, 360), "white")
    draw = ImageDraw.Draw(image)
    draw.text((40, 30), "Product Overview", fill=(0, 0, 0))
    draw.text((40, 62), "This page explains the main product.", fill=(0, 0, 0))
    draw.rectangle((300, 70, 455, 210), fill=(42, 120, 210))
    draw.text((40, 245), "The chart below is important.", fill=(0, 0, 0))
    image.save(image_path)

    layout = {
        "blocks": [
            {
                "type": "text",
                "order": 1,
                "text": "Product Overview",
                "bbox": [40, 25, 210, 50],
                "role": "heading",
            },
            {
                "type": "text",
                "order": 2,
                "text": "This page explains the main product.",
                "bbox": [40, 58, 275, 82],
                "role": "paragraph",
            },
            {
                "type": "image",
                "order": 3,
                "bbox": [292, 62, 464, 218],
                "asset_type": "product",
                "importance": "high",
                "reason": "Main product image",
                "caption": "Product image",
            },
            {
                "type": "image",
                "order": 4,
                "bbox": [20, 310, 38, 328],
                "asset_type": "icon",
                "importance": "low",
                "reason": "decorative icon",
                "caption": "",
            },
        ]
    }

    output_docx = tmp_path / "web.docx"
    result = render_web_asset_preserve_docx_from_layout(
        image_path=image_path,
        layout=layout,
        output_docx_path=output_docx,
        assets_dir=tmp_path / "assets",
    )

    assert output_docx.exists()
    assert result.asset_count == 1
    assert "Product Overview" in result.raw_text
    assert "Main product image" in result.raw_text
    assert result.layout["render"]["debug_overlays"]
    assert Path(result.layout["render"]["debug_overlays"][0]).exists()

    document = Document(output_docx)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    assert "Product Overview" in text
    assert "This page explains the main product." in text

    with zipfile.ZipFile(output_docx) as docx_zip:
        media_files = [name for name in docx_zip.namelist() if name.startswith("word/media/")]
    assert len(media_files) >= 1


def test_render_web_asset_preserve_refines_sloppy_image_box(tmp_path):
    image_path = tmp_path / "web_sloppy.png"
    image = Image.new("RGB", (520, 360), "white")
    draw = ImageDraw.Draw(image)
    draw.text((40, 30), "Dashboard", fill=(0, 0, 0))
    draw.rectangle((180, 96, 352, 210), fill=(40, 160, 110))
    image.save(image_path)

    layout = {
        "blocks": [
            {
                "type": "text",
                "order": 1,
                "text": "Dashboard",
                "bbox": [40, 25, 140, 52],
                "role": "heading",
            },
            {
                "type": "image",
                "order": 2,
                "bbox": [150, 70, 382, 238],
                "asset_type": "chart",
                "importance": "high",
                "reason": "Important chart",
                "caption": "",
            },
        ]
    }

    result = render_web_asset_preserve_docx_from_layout(
        image_path=image_path,
        layout=layout,
        output_docx_path=tmp_path / "web_sloppy.docx",
        assets_dir=tmp_path / "assets_sloppy",
    )

    block = result.layout["pages"][0]["layout"]["blocks"][1]
    refined = block["bbox_refined"]

    assert refined != [150, 70, 382, 238]
    assert refined[0] >= 170
    assert refined[1] >= 86
    assert refined[2] <= 362
    assert refined[3] <= 220
