#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""PrepararPlantilla.py - Analiza una plantilla .docx y genera el mapa de estilos en Settings.json."""


import sys
import os
import json
import argparse
from docx import Document
from docx.enum.style import WD_STYLE_TYPE


# Canonical style roles
ROLE_MAP_DEFAULTS = {
    "H1": ["heading 1", "titulo 1", "título 1"],
    "H2": ["heading 2", "titulo 2", "título 2"],
    "H3": ["heading 3", "titulo 3", "título 3"],
    "Body": ["normal", "body text", "cuerpo"],
    "Bullet": ["list bullet", "bullet", "viñeta"],
    "Code": ["code", "código", "source code", "mono"],
    "Caption": ["caption", "leyenda", "pie de imagen"],
    "Footer": ["footer", "pie de página"],
}

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Settings.json")


def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"General": {}, "Python": {}, "Styles": {}}


def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2, ensure_ascii=False)


def type_label(style):
    mapping = {
        WD_STYLE_TYPE.PARAGRAPH: "paragraph",
        WD_STYLE_TYPE.CHARACTER: "character",
        WD_STYLE_TYPE.TABLE: "table",
        WD_STYLE_TYPE.LIST: "list",
    }
    return mapping.get(style.type, "unknown")


def find_best_match(styles, role, candidates):
    """Find the best style name match for a given role."""
    for style in styles:
        if style.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        name_lower = style.name.lower()
        for candidate in candidates:
            if candidate in name_lower:
                return style.name
    return None


def analyze_template(template_path, debug_style=None):
    doc = Document(template_path)
    styles = list(doc.styles)

    print("=" * 60)
    print("ESTILOS ENCONTRADOS EN LA PLANTILLA")
    print("=" * 60)
    print(f"{'Nombre':<35} {'Tipo':<12} {'ID interno'}")
    print("-" * 60)

    for style in sorted(styles, key=lambda s: (type_label(s), s.name)):
        print(f"{style.name:<35} {type_label(style):<12} {style.style_id}")

    if debug_style:
        print("\n" + "=" * 60)
        print(f"DEBUG XML - Estilo: {debug_style}")
        print("=" * 60)
        for style in styles:
            if style.name.lower() == debug_style.lower() or style.style_id.lower() == debug_style.lower():
                from lxml import etree
                print(etree.tostring(style.element, pretty_print=True).decode("utf-8"))
                break
        else:
            print(f"[AVISO] Estilo '{debug_style}' no encontrado.")

    # Auto-map
    style_map = {}
    for role, candidates in ROLE_MAP_DEFAULTS.items():
        match = find_best_match(styles, role, candidates)
        if match:
            style_map[role] = match
        else:
            # Fallback defaults
            defaults = {"H1": "Heading 1", "H2": "Heading 2", "H3": "Heading 3",
                        "Body": "Normal", "Bullet": "List Bullet", "Code": "Code",
                        "Caption": "Caption", "Footer": "Footer"}
            style_map[role] = defaults.get(role, "Normal")

    print("\n" + "=" * 60)
    print("MAPEO DE ESTILOS RESULTANTE")
    print("=" * 60)
    for role, name in style_map.items():
        print(f"  {role:<10} -> {name}")

    # Check for cover image placeholder
    has_cover_image = False
    for para in doc.paragraphs:
        for run in para.runs:
            if run._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'):
                has_cover_image = True
                break
        if has_cover_image:
            break

    if has_cover_image:
        print("\n[INFO] La plantilla contiene una imagen en la primera página (posible placeholder de portada).")

    # Save
    settings = load_settings()
    settings["Styles"] = style_map
    save_settings(settings)
    print(f"\n[OK] Settings.json actualizado con el mapa de estilos.")
    return style_map


def main():
    parser = argparse.ArgumentParser(description="Analiza plantilla .docx y mapea estilos.")
    parser.add_argument("template", help="Ruta a la plantilla .docx")
    parser.add_argument("--debug", metavar="ESTILO", help="Muestra el XML interno del estilo indicado")
    args = parser.parse_args()

    if not os.path.exists(args.template):
        print(f"[ERROR] No se encuentra: {args.template}")
        sys.exit(1)

    analyze_template(args.template, args.debug)


if __name__ == "__main__":
    main()
