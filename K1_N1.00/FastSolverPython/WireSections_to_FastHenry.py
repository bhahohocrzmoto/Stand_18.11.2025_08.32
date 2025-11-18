#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WireSections_to_FastHenry_fromMacro.py
---------------------------------------

Standalone Python script (no FreeCAD) that reads a "Wire_Sections.txt"
in the SAME format expected by Build_FH_from_WireSections.FCMacro and
writes a FastHenry2-compatible input file.

- Each Section becomes:
    - A chain of segments between consecutive points.
    - One .external port between FIRST and LAST node of that section.

- Trace width/height:
    - Global defaults: DEFAULT_SEG_WIDTH, DEFAULT_SEG_HEIGHT
    - Optional per-section overrides: SECTION_WH[section_name] = (w, h)

Usage:
    python WireSections_to_FastHenry_fromMacro.py Wire_Sections.txt \
        --trace-width 0.25 --trace-thickness 0.035 --fmin 1e3 --fmax 1e3

Output:
    If not specified, "Wire_Sections.inp" next to the input file.

This script does NOT depend on FreeCAD or EM; it's pure Python.
"""

import argparse
from pathlib import Path

# --------------------------------------------------------------------------- #
# --------------------------- CONFIGURATION --------------------------------- #
# --------------------------------------------------------------------------- #

# Default cross-section dimensions (same *idea* as DEFAULT_SEG_WIDTH_MM, etc.).
# These are in the SAME length units as the Wire_Sections file header:
#   - If header is "mm": values are in mm
#   - If header is "cm": values are in cm
DEFAULT_SEG_WIDTH  = 0.25   # e.g. 0.25 mm
DEFAULT_SEG_HEIGHT = 0.035  # e.g. 0.035 mm (35 µm copper)

# Optional per-section overrides, like your SECTION_WH_MM:
# Example:
# SECTION_WH = {
#     "Section-1": (0.25, 0.035),   # width, height
#     "Section-2": (0.30, 0.035),
# }
SECTION_WH = {}


# --------------------------------------------------------------------------- #
# ------------------------------ PARSER ------------------------------------- #
# --------------------------------------------------------------------------- #

def parse_wire_sections(txt_path):
    """
    Parse Wire_Sections.txt in the same spirit as Build_FH_from_WireSections:

    Expected structure (non-empty lines):
        1) units token, e.g. "mm" or "cm"
        2) parameter line (vol_res_cm=..., etc.) → ignored
        3+) "Section-Name, X, Y, Z, scalar"

    Returns
    -------
    units : str
        "MM" or "CM" (normalized to uppercase, defaults to "MM" if unknown).
    sections : dict[str, list[tuple[int,float,float,float,int]]]
        Maps section name to list of points:
            (idx_in_section, x, y, z, src_line_number)
    """
    sections = {}
    units = "MM"  # default

    lines = []
    with open(txt_path, "r", encoding="utf-8") as f:
        for ln in f:
            ln = ln.strip()
            if ln:
                lines.append(ln)

    if not lines:
        raise ValueError("Input file is empty or only whitespace.")

    # First non-empty line is usually the units token ("mm" or "cm")
    first = lines[0].strip().lower()
    if first in ("mm", "millimeter", "millimetre"):
        units = "MM"
        start_idx = 2  # skip parameter line as well
    elif first in ("cm", "centimeter", "centimetre"):
        units = "CM"
        start_idx = 2
    else:
        # If it's not obviously units, we treat it as data
        units = "MM"
        start_idx = 0

    # Safety: if the file is too short, don't skip too many lines
    if start_idx >= len(lines):
        start_idx = 0

    # Now parse Section lines starting from start_idx
    line_number = 1  # human-friendly line numbers
    for idx_in_list, line in enumerate(lines):
        # We want the real line number as seen by user -> line_number
        # (here just sequential in the filtered list; if you want absolute line
        # in the original file with blanks, you'd need a different approach.)
        if idx_in_list < start_idx:
            line_number += 1
            continue

        parts = [p.strip() for p in line.split(",")]
        if len(parts) < 4:
            line_number += 1
            continue

        sec_name = parts[0]
        if not sec_name.startswith("Section-"):
            line_number += 1
            continue

        try:
            x = float(parts[1])
            y = float(parts[2])
            z = float(parts[3])
        except ValueError:
            line_number += 1
            continue

        if sec_name not in sections:
            sections[sec_name] = []

        # idx_in_section is assigned later; store placeholder (None for now).
        sections[sec_name].append([None, x, y, z, line_number])

        line_number += 1

    # Assign per-section indices (1-based), like the macro does
    for sec_name, pts in sections.items():
        for k, row in enumerate(pts, start=1):
            row[0] = k  # idx within this section

    return units, sections


# --------------------------------------------------------------------------- #
# ------------------------------ HELPERS ------------------------------------ #
# --------------------------------------------------------------------------- #

def units_to_sigma(units):
    """
    Pick a reasonable default copper conductivity in FastHenry units.

    Copper: sigma_SI ~ 5.8e7 S/m

    If lengths in:
      - M : sigma ≈ 5.8e7
      - CM: sigma ≈ 5.8e5
      - MM: sigma ≈ 5.8e4

    This keeps ohmic resistance roughly correct when you input geometry
    in those length units.
    """
    units = units.upper()
    if units == "M":
        return 5.8e7
    elif units == "CM":
        return 5.8e5
    elif units == "MM":
        return 5.8e4
    else:
        # default to mm-scaling
        return 5.8e4


def section_sort_key(sec_name):
    """
    Sort key so 'Section-1', 'Section-2', ..., 'Section-10' are in numeric order.

    Returns
    -------
    tuple
        (base, number) if '-' present and number parses, else (base, sec_name)
    """
    name = sec_name.strip()
    if "-" in name:
        base, num = name.rsplit("-", 1)
        base = base.strip()
        num = num.strip()
        try:
            return (base, int(num))
        except ValueError:
            return (base, name)
    return ("", name)


def make_node_prefix(sec_name):
    """
    Convert a section name like 'Section-3' into a compact prefix like 'S3'
    for naming nodes in FastHenry.

    If the pattern is different, we fall back to a cleaned-up section name.
    """
    sec = sec_name.strip()
    if "-" in sec:
        base, num = sec.rsplit("-", 1)
        base = base.strip()
        num = num.strip()
        if base.lower().startswith("section"):
            return f"S{num}"
        else:
            return (base + "_" + num).replace(" ", "_")
    else:
        return sec.replace(" ", "_")


def make_node_name(sec_name, idx):
    """Return a FastHenry node label similar to FreeCAD's convention."""

    safe_section = sec_name.strip().replace(" ", "_")
    return f"N{safe_section}_Node_{idx}"


def format_coord(value, force_decimal=False):
    """Format coordinates like FreeCAD's EM workbench output."""

    if abs(value) < 1e-12:
        value = 0.0
    text = f"{value:.8f}".rstrip("0").rstrip(".")
    if not text:
        text = "0"
    if force_decimal and "." not in text:
        text += ".0"
    return text


# --------------------------------------------------------------------------- #
# ------------------------ FASTHENRY WRITER --------------------------------- #
# --------------------------------------------------------------------------- #

def write_fasthenry_input(
    out_path,
    units,
    sections,
    default_width,
    default_height,
    sigma=None,
    freq_min=1.0,
    freq_max=1e9,
    freq_decades=1.0,
    nhinc=1,
    nwinc=1,
    rh=2,
    rw=2,
):
    """
    Write a FastHenry2-compatible input file.

    Parameters
    ----------
    out_path : str or Path
        Output file path for .inp/.txt.
    units : str
        "MM", "CM", "M" etc. (we only handle MM/CM/M for sigma scaling).
    sections : dict[str, list[list[idx,x,y,z,line_no]]]
        Parsed sections from parse_wire_sections().
    default_width, default_height : float
        Global default cross-section (same units as coordinates).
    sigma : float or None
        Conductivity; if None, we use units_to_sigma(units).
    freq_min, freq_max : float
        Frequency sweep for .freq card (Hz).
    freq_decades : float
        Number of points-per-decade for .freq (FastHenry's ndec parameter).
    nhinc, nwinc : int
        Number of subdivisions for the height/width directions.
    rh, rw : int
        Aspect ratio hints (FastHenry parameters rh/rw).
    """
    out_path = Path(out_path)
    units = units.upper()

    if sigma is None:
        sigma = units_to_sigma(units)

    # We'll collect formatted nodes, segments and ports so that we can emit
    # blocks identical to FreeCAD's FastHenry exporter.
    all_nodes = []
    segments = []
    ports = []

    for sec_name in sorted(sections.keys(), key=section_sort_key):
        pts = sections[sec_name]
        if len(pts) < 2:
            continue

        w_sec, h_sec = SECTION_WH.get(sec_name, (default_width, default_height))
        node_names = []
        for idx, x, y, z, _line_no in pts:
            node_name = make_node_name(sec_name, idx)
            node_names.append(node_name)
            all_nodes.append((node_name, x, y, z))

        for seg_idx in range(len(node_names) - 1):
            segments.append((node_names[seg_idx], node_names[seg_idx + 1], w_sec, h_sec))

        ports.append((node_names[0], node_names[-1]))

    with out_path.open("w", encoding="utf-8", newline="") as f:
        def write_line(text=""):
            """Emit a CRLF-terminated line to match FreeCAD's deck byte-for-byte."""

            f.write(text + "\r\n")

        write_line("* FastHenry input file created using FreeCAD's ElectroMagnetic Workbench")
        write_line("* See http://www.freecad.org, http://www.fastfieldsolvers.com and http://epc-co.com")
        write_line()
        write_line(f".units {units.lower()}")
        write_line()
        write_line(
            f".default sigma={format_coord(sigma, force_decimal=True)} nhinc={nhinc} "
            f"nwinc={nwinc} rh={rh} rw={rw}"
        )
        write_line()
        write_line("* Nodes")
        for node_name, x, y, z in all_nodes:
            write_line(
                f"{node_name} x={format_coord(x, True)} y={format_coord(y, True)} "
                f"z={format_coord(z, True)}"
            )

        write_line()
        write_line("* Segments")
        for idx, (n1, n2, w_val, h_val) in enumerate(segments):
            if idx == 0:
                elem_name = "EFHSegment"
            else:
                elem_name = f"EFHSegment{idx:03d}"
            write_line(f"{elem_name} {n1} {n2} w={format_coord(w_val)} h={format_coord(h_val)}")

        write_line()
        write_line("* Ports")
        for n_start, n_end in ports:
            write_line(f".external {n_start} {n_end}")

        write_line()
        write_line(
            f".freq fmin={format_coord(freq_min, True)} "
            f"fmax={format_coord(freq_max, True)} ndec={format_coord(freq_decades, True)}"
        )
        write_line()
        write_line(".end")

    return {
        "sections": len(ports),
        "nodes": len(all_nodes),
        "segments": len(segments),
        "ports": len(ports),
    }


# --------------------------------------------------------------------------- #
# ----------------------------- CLI ENTRYPOINT ------------------------------ #
# --------------------------------------------------------------------------- #

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Convert Wire_Sections.txt (as used by Build_FH_from_WireSections.FCMacro) "
            "into a FastHenry2 input file."
        )
    )
    parser.add_argument("infile", help="Input Wire_Sections.txt")
    parser.add_argument(
        "-o",
        "--outfile",
        help="Output FastHenry2 file (.inp/.txt). "
             "Default: same name as input, with .inp extension.",
    )
    parser.add_argument(
        "--trace-width",
        type=float,
        default=DEFAULT_SEG_WIDTH,
        help=(
            "Global default trace width in the same units as the Wire_Sections file "
            f"(default: {DEFAULT_SEG_WIDTH})"
        ),
    )
    parser.add_argument(
        "--trace-thickness",
        type=float,
        default=DEFAULT_SEG_HEIGHT,
        help=(
            "Global default trace thickness in the same units as the Wire_Sections file "
            f"(default: {DEFAULT_SEG_HEIGHT})"
        ),
    )
    parser.add_argument(
        "--sigma",
        type=float,
        default=None,
        help=(
            "Conductivity for .Default sigma (1/(ohm*unit)). "
            "If not given, a copper-like default is chosen based on units."
        ),
    )
    parser.add_argument(
        "--fmin",
        type=float,
        default=1.0,
        help="Minimum frequency in Hz for .freq (default: 1.0)",
    )
    parser.add_argument(
        "--fmax",
        type=float,
        default=1e9,
        help="Maximum frequency in Hz for .freq (default: 1e9)",
    )
    parser.add_argument(
        "--freq-decades",
        type=float,
        default=1.0,
        help="Points per decade for .freq ndec parameter (default: 1.0)",
    )
    parser.add_argument(
        "--nhinc",
        type=int,
        default=1,
        help="Number of subdivisions along trace thickness (FastHenry nhinc)",
    )
    parser.add_argument(
        "--nwinc",
        type=int,
        default=1,
        help="Number of subdivisions along trace width (FastHenry nwinc)",
    )
    parser.add_argument(
        "--rh",
        type=int,
        default=2,
        help="Aspect-ratio control rh passed to .default (default: 2)",
    )
    parser.add_argument(
        "--rw",
        type=int,
        default=2,
        help="Aspect-ratio control rw passed to .default (default: 2)",
    )

    args = parser.parse_args()

    in_path = Path(args.infile)
    if not in_path.is_file():
        raise SystemExit(f"Input file not found: {in_path}")

    # Parse the Wire_Sections file using macro-compatible logic
    units, sections = parse_wire_sections(in_path)
    if not sections:
        raise SystemExit("No Section-* data found in input file.")

    # Determine output path
    if args.outfile:
        out_path = Path(args.outfile)
    else:
        out_path = in_path.with_suffix(".inp")

    # Write FastHenry2 input
    summary = write_fasthenry_input(
        out_path=out_path,
        units=units,
        sections=sections,
        default_width=args.trace_width,
        default_height=args.trace_thickness,
        sigma=args.sigma,
        freq_min=args.fmin,
        freq_max=args.fmax,
        freq_decades=args.freq_decades,
        nhinc=args.nhinc,
        nwinc=args.nwinc,
        rh=args.rh,
        rw=args.rw,
    )

    abs_out = out_path.resolve()
    print(
        "[OK] Written FastHenry2 input to:"
        f" {abs_out} (sections={summary['sections']}, nodes={summary['nodes']}, "
        f"segments={summary['segments']}, ports={summary['ports']})"
    )


if __name__ == "__main__":
    main()
