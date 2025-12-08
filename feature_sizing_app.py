from __future__ import annotations

import json
import math
import os
import re
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


DEFAULT_SIZE_TO_HOURS = {"XS": 2, "S": 4, "M": 8, "L": 16, "XL": 24}
SIZE_ORDER = ["XS", "S", "M", "L", "XL"]


@dataclass
class Requirement:
    text: str
    size: str
    notes: str = "—"


@dataclass
class ScreenDetail:
    title: str
    description: str


@dataclass
class ModuleSpec:
    name: str
    description: str
    requirements: List[Requirement] = field(default_factory=list)
    summary: str = ""
    screen_details: List[ScreenDetail] = field(default_factory=list)
    screen_count: int | None = None

    def total_hours(self, size_to_hours: Dict[str, float]) -> float:
        return sum(size_to_hours.get(req.size.upper(), 0) for req in self.requirements)

    def resolved_screen_count(self) -> int:
        if self.screen_details:
            return len(self.screen_details)
        if self.screen_count is not None:
            return max(1, self.screen_count)
        return max(1, math.ceil(len(self.requirements) / 2))


def ordered_size_keys(size_hours: Dict[str, float]) -> List[str]:
    primary = [size for size in SIZE_ORDER if size in size_hours]
    extras = [size for size in size_hours if size not in SIZE_ORDER]
    combined = primary + extras
    return combined or SIZE_ORDER.copy()


def get_openai_client(api_key: str | None = None) -> OpenAI:
    key = api_key or os.getenv("OPENAI_API_KEY")
    if not key:
        raise RuntimeError(
            "Missing OpenAI API key. Set the OPENAI_API_KEY environment variable "
            "or provide it in the UI."
        )
    return OpenAI(api_key=key)


def parse_llm_response(payload: str) -> Tuple[List[ModuleSpec], List[str], List[str], Dict[str, float]]:
    try:
        data = json.loads(payload)
    except json.JSONDecodeError as error:
        raise ValueError(f"LLM response was not valid JSON: {error}") from error

    size_hours_raw = data.get("size_hours") or DEFAULT_SIZE_TO_HOURS
    size_hours: Dict[str, float] = {}
    for size, hours in size_hours_raw.items():
        if size and isinstance(hours, (int, float)):
            size_hours[size.upper()] = float(hours)
    if not size_hours:
        size_hours = DEFAULT_SIZE_TO_HOURS.copy()
    else:
        for size, hours in DEFAULT_SIZE_TO_HOURS.items():
            size_hours.setdefault(size, float(hours))

    modules_data = data.get("modules", [])
    modules: List[ModuleSpec] = []
    for module in modules_data:
        requirements = []
        for req in module.get("requirements", []):
            requirements.append(
                Requirement(
                    text=req.get("text", "").strip() or "Requirement TBD",
                    size=req.get("size", "M").upper(),
                    notes=req.get("notes", "—") or "—",
                )
            )
        screen_details: List[ScreenDetail] = []
        for screen in module.get("screens", []):
            screen_details.append(
                ScreenDetail(
                    title=(screen.get("title") or "Screen").strip(),
                    description=(screen.get("description") or "Description TBD").strip(),
                )
            )
        screen_count_raw = module.get("screen_count")
        try:
            screen_count = int(screen_count_raw) if screen_count_raw is not None else None
        except (TypeError, ValueError):
            screen_count = None
        modules.append(
            ModuleSpec(
                name=module.get("name", "Module").strip() or "Module",
                description=module.get("description", "Description TBD").strip(),
                requirements=requirements,
                summary=module.get("summary", "").strip(),
                screen_details=screen_details,
                screen_count=screen_count,
            )
        )

    risks = data.get("risks", [])
    questions = data.get("questions", [])
    return modules, risks, questions, size_hours


def request_plan_from_llm(description: str, api_key: str | None = None):
    client = get_openai_client(api_key)
    system_prompt = (
        "You are a senior presales analyst. Produce a feature sizing plan as STRICT JSON "
        "using this schema:\n"
        "{\n"
        '  "size_hours": {"XS": 2, "S": 4, "M": 8, "L": 16, "XL": 24},\n'
        '  "modules": [\n'
        "    {\n"
        '      "name": "string",\n'
        '      "description": "string",\n'
        '      "summary": "string",\n'
        '      "screen_count": 4,\n'
        '      "screens": [\n'
        '         {"title": "Screen title", "description": "purpose"}\n'
        "      ],\n"
        '      "requirements": [\n'
        '         {"text": "one-line, very granular", "size": "M", "notes": ""}\n'
        "      ]\n"
        "    }\n"
        "  ],\n"
        '  "risks": ["risk 1", "..."],\n'
        '  "questions": ["question 1", "..."]\n'
        "}\n"
        "Rules:\n"
        "- Always include XS/S/M/L/XL keys in size_hours (uppercase) and map to estimated dev hours.\n"
        "- Provide 7-12 granular requirements per module, each with size XS/S/M/L/XL only.\n"
        "- Each module must list every planned screen with title + short purpose sentence.\n"
        "- Keep prose concise; no markdown or commentary outside the JSON."
    )
    user_prompt = (
        "Feature description:\n"
        f"{description.strip()}\n\n"
        "Respond with the JSON structure described in the system prompt. Do not include any prose."
    )
    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )
    content = response.choices[0].message.content
    if not content:
        raise RuntimeError("Empty response from language model.")
    return parse_llm_response(content)


def collect_feature_description() -> str:
    print(
        "Paste the free-text feature description. Press Enter twice when you're done:\n"
    )
    lines: List[str] = []
    blank_streak = 0
    while True:
        try:
            line = input()
        except EOFError:
            break

        if not line.strip():
            blank_streak += 1
            if blank_streak >= 2:
                break
        else:
            blank_streak = 0

        lines.append(line)

    description = "\n".join(lines).strip()
    if not description:
        raise ValueError("Feature description cannot be empty.")
    return description


def sanitize_sheet_name(name: str, existing: set) -> str:
    sanitized = re.sub(r"[\\/*?:\[\]]", "", name)
    if len(sanitized) > 31:
        sanitized = sanitized[:31]
    base = sanitized or "Module"
    final_name = base
    suffix = 1
    while final_name.lower() in {n.lower() for n in existing}:
        trimmed = base[: 31 - len(str(suffix)) - 1]
        final_name = f"{trimmed}_{suffix}"
        suffix += 1
    existing.add(final_name)
    return final_name


def auto_fit_columns(ws, min_width: int = 12):
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
        length = max(length + 2, min_width)
        ws.column_dimensions[get_column_letter(column_cells[0].column)].width = min(length, 60)


def write_module_sheet(
    wb: Workbook,
    module: ModuleSpec,
    existing_sheet_names: set,
    size_range_ref: str,
    hours_range_ref: str,
) -> Tuple[str, str]:
    sheet_name = sanitize_sheet_name(module.name, existing_sheet_names)
    ws = wb.create_sheet(title=sheet_name)
    bold = Font(bold=True)

    ws["A1"].value = "Module Name"
    ws["A1"].font = bold
    ws["B1"].value = module.name

    ws["A2"].value = "Module Description"
    ws["A2"].font = bold
    ws["B2"].value = module.description
    ws["B2"].alignment = Alignment(wrap_text=True)

    ws["A3"].value = "Number of Screens"
    ws["A3"].font = bold
    ws["B3"].value = module.resolved_screen_count()

    header_row = 5
    ws[f"A{header_row}"].value = "Requirement"
    ws[f"B{header_row}"].value = "Size (XS/S/M/L/XL)"
    ws[f"C{header_row}"].value = "Notes"
    for col in "ABC":
        ws[f"{col}{header_row}"].font = bold

    req_start = header_row + 1
    if not module.requirements:
        module.requirements.append(
            Requirement(
                text="Placeholder requirement derived from description.",
                size="M",
            )
        )

    for idx, req in enumerate(module.requirements):
        row = req_start + idx
        ws[f"A{row}"] = req.text
        ws[f"B{row}"] = req.size
        ws[f"C{row}"] = req.notes
        ws[f"A{row}"].alignment = Alignment(wrap_text=True)

    req_end = req_start + len(module.requirements) - 1
    screens_label_row = req_end + 2
    ws[f"A{screens_label_row}"].value = "Screens"
    ws[f"A{screens_label_row}"].font = bold

    screen_header_row = screens_label_row + 1
    ws[f"A{screen_header_row}"].value = "Screen Title"
    ws[f"B{screen_header_row}"].value = "Description"
    ws[f"A{screen_header_row}"].font = bold
    ws[f"B{screen_header_row}"].font = bold

    screen_rows = module.screen_details or [
        ScreenDetail(
            title="Placeholder Screen",
            description="High-level view derived from requirements.",
        )
    ]
    for idx, screen in enumerate(screen_rows):
        row = screen_header_row + 1 + idx
        ws[f"A{row}"] = screen.title
        ws[f"B{row}"] = screen.description
        ws[f"B{row}"].alignment = Alignment(wrap_text=True)

    screen_table_end = screen_header_row + len(screen_rows)
    summary_row = screen_table_end + 2

    ws[f"A{summary_row}"].value = "Module Summary:"
    ws[f"A{summary_row}"].font = bold
    ws[f"B{summary_row}"].value = module.summary
    ws[f"B{summary_row}"].alignment = Alignment(wrap_text=True)

    ws[f"A{summary_row + 1}"].value = "Screens:"
    ws[f"A{summary_row + 1}"].font = bold
    ws[f"B{summary_row + 1}"].value = module.resolved_screen_count()

    ws[f"A{summary_row + 2}"].value = "Total Module Hours (calculated):"
    ws[f"A{summary_row + 2}"].font = bold
    total_cell_ref = f"B{summary_row + 2}"
    count_range_ref = f"$B${req_start}:$B${req_end}"
    ws[total_cell_ref] = (
        f"=SUMPRODUCT(COUNTIF({count_range_ref},{size_range_ref}),{hours_range_ref})"
    )

    auto_fit_columns(ws)
    return sheet_name, total_cell_ref


def build_resource_loading_sheet(
    wb: Workbook,
    module_totals: List[Tuple[str, str, str]],
    size_to_hours: Dict[str, float],
    size_order: List[str],
):
    ws = wb.create_sheet(title="ResourceLoading", index=0)
    bold = Font(bold=True)

    ws["A1"].value = "T-shirt Sizing to Hours"
    ws["A1"].font = bold

    ws["A2"].value = "Size"
    ws["B2"].value = "Hours (editable)"
    ws["A2"].font = bold
    ws["B2"].font = bold

    size_start_row = 3
    for offset, size in enumerate(size_order):
        row = size_start_row + offset
        ws[f"A{row}"] = size
        ws[f"B{row}"] = size_to_hours.get(size, DEFAULT_SIZE_TO_HOURS.get(size, 0))

    size_end_row = size_start_row + len(size_order) - 1
    alloc_header_row = size_end_row + 2
    ws[f"A{alloc_header_row}"].value = "Effort Allocation"
    ws[f"A{alloc_header_row}"].font = bold
    ws[f"C{alloc_header_row}"].value = "Ratio (editable)"
    ws[f"C{alloc_header_row}"].font = bold

    total_refs = ",".join([f"'{sheet}'!{cell}" for sheet, _, cell in module_totals])
    dev_row = alloc_header_row + 1
    ws[f"A{dev_row}"] = "Dev Hours"
    ws[f"A{dev_row}"].font = bold
    ws[f"B{dev_row}"] = f"=SUM({total_refs})" if total_refs else 0

    allocations = [
        ("QA Hours", 0.4),
        ("PM Hours", 0.15),
        ("Architect Hours", 0.10),
        ("Buffer Hours", 0.20),
    ]
    allocation_rows: Dict[str, int] = {}
    for index, (label, ratio) in enumerate(allocations, start=1):
        row = dev_row + index
        allocation_rows[label] = row
        ws[f"A{row}"] = label
        ws[f"A{row}"].font = bold
        ws[f"C{row}"] = ratio
        ws[f"B{row}"] = f"=$B${dev_row}*$C{row}"

    hours_per_day_row = dev_row + len(allocations) + 2
    ws[f"A{hours_per_day_row}"] = "Hours per Resource Day"
    ws[f"A{hours_per_day_row}"].font = bold
    ws[f"B{hours_per_day_row}"] = 8

    util_header_row = hours_per_day_row + 1
    ws[f"A{util_header_row}"].value = "Resource Utilization"
    ws[f"A{util_header_row}"].font = bold
    ws[f"A{util_header_row + 1}"].value = "Role"
    ws[f"B{util_header_row + 1}"].value = "Hours"
    ws[f"C{util_header_row + 1}"].value = "Resource Days"
    for col in "ABC":
        ws[f"{col}{util_header_row + 1}"].font = bold

    util_rows = [
        ("Dev", f"$B${dev_row}"),
        ("QA", f"$B${allocation_rows['QA Hours']}"),
        ("PM", f"$B${allocation_rows['PM Hours']}"),
        ("Architect", f"$B${allocation_rows['Architect Hours']}"),
        ("Buffer", f"$B${allocation_rows['Buffer Hours']}"),
    ]
    for idx, (role, hours_cell) in enumerate(util_rows, start=util_header_row + 2):
        ws[f"A{idx}"] = role
        ws[f"B{idx}"] = f"={hours_cell}"
        ws[f"C{idx}"] = f"={hours_cell}/$B${hours_per_day_row}"

    auto_fit_columns(ws, min_width=14)


def build_workbook(
    modules: List[ModuleSpec],
    size_to_hours: Dict[str, float],
    output_path: Path,
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    size_order = ordered_size_keys(size_to_hours)
    size_start_row = 3
    size_end_row = size_start_row + len(size_order) - 1
    size_range_ref = f"ResourceLoading!$A${size_start_row}:$A${size_end_row}"
    hours_range_ref = f"ResourceLoading!$B${size_start_row}:$B${size_end_row}"

    module_totals: List[Tuple[str, str, str]] = []
    existing_sheet_names: set = set()

    for module in modules:
        sheet_name, total_cell = write_module_sheet(
            wb=wb,
            module=module,
            existing_sheet_names=existing_sheet_names,
            size_range_ref=size_range_ref,
            hours_range_ref=hours_range_ref,
        )
        module_totals.append((sheet_name, module.name, total_cell))

    build_resource_loading_sheet(
        wb=wb,
        module_totals=module_totals,
        size_to_hours=size_to_hours,
        size_order=size_order,
    )
    wb.active = wb["ResourceLoading"]
    wb.save(output_path)


def format_preview(
    modules: List[ModuleSpec],
    size_to_hours: Dict[str, float],
    risks: List[str],
    questions: List[str],
) -> str:
    dev_hours = sum(module.total_hours(size_to_hours) for module in modules)
    qa_hours = round(dev_hours * 0.4, 2)
    pm_hours = round(dev_hours * 0.15, 2)
    architect_hours = round(dev_hours * 0.1, 2)
    buffer_hours = round(dev_hours * 0.2, 2)

    lines: List[str] = []
    for module in modules:
        lines.append(f"Module: {module.name}")
        lines.append(f"Description: {module.description}")
        lines.append(f"Screens: {module.resolved_screen_count()}")
        if module.screen_details:
            lines.append("Screen Details:")
            for screen in module.screen_details:
                lines.append(f"  - {screen.title}: {screen.description}")
        else:
            lines.append("Screen Details: (TBD)")
        lines.append("Requirements:")
        for req in module.requirements:
            notes = req.notes if req.notes else "—"
            lines.append(f"  - [{req.size}] {req.text} ({notes})")
        lines.append(f"Module Summary: {module.summary}")
        lines.append(f"Total Module Hours: {module.total_hours(size_to_hours)}")
        lines.append("")

    lines.append("Resource Loading:")
    lines.append(
        "  Size to Hours Mapping: "
        + ", ".join(f"{k}={v}" for k, v in size_to_hours.items())
    )
    lines.append(f"  Dev Hours: {dev_hours}")
    lines.append(f"  QA Hours (40%): {qa_hours}")
    lines.append(f"  PM Hours (15%): {pm_hours}")
    lines.append(f"  Architect Hours (10%): {architect_hours}")
    lines.append(f"  Buffer (20%): {buffer_hours}")
    lines.append("")

    lines.append("Risks / Missing Information:")
    for risk in risks:
        lines.append(f"  - {risk}")
    lines.append("")

    lines.append("Clarification Questions:")
    for question in questions:
        lines.append(f"  - {question}")

    return "\n".join(lines)


def generate_analysis(description: str, output_folder: Path, api_key: str | None = None):
    modules, risks, questions, size_to_hours = request_plan_from_llm(description, api_key)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_folder.mkdir(parents=True, exist_ok=True)
    output_path = output_folder / f"feature_sizing_{timestamp}.xlsx"
    build_workbook(modules, size_to_hours, output_path)
    preview = format_preview(modules, size_to_hours, risks, questions)
    return output_path, preview


def main():
    description = collect_feature_description()
    output_path, preview = generate_analysis(description, Path.cwd())
    divider = "=" * 60
    print(f"\n{divider}\nTEXTUAL PREVIEW\n{divider}")
    print(preview)
    print(f"\nExcel workbook saved to: {output_path.resolve()}\n")


def run_gui():
    root = tk.Tk()
    root.title("Feature Sizing Assistant")
    root.geometry("900x700")

    output_dir = tk.StringVar(value=str(Path.cwd()))
    api_key_var = tk.StringVar(value=os.getenv("OPENAI_API_KEY", ""))

    def choose_output_dir():
        directory = filedialog.askdirectory(initialdir=output_dir.get() or str(Path.cwd()))
        if directory:
            output_dir.set(directory)

    def generate():
        text = description_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showerror("Missing description", "Please paste a feature description first.")
            return

        try:
            out_path, preview = generate_analysis(
                text,
                Path(output_dir.get() or Path.cwd()),
                api_key=api_key_var.get().strip() or None,
            )

            preview_text.configure(state="normal")
            preview_text.delete("1.0", tk.END)
            preview_text.insert(tk.END, preview)
            preview_text.configure(state="disabled")

            status_var.set(f"Excel workbook saved to: {out_path}")
        except Exception as exc:  # pragma: no cover - UI safeguard
            messagebox.showerror("Error", f"Failed to generate workbook:\n{exc}")

    def open_output_folder():
        try:
            path = Path(output_dir.get() or Path.cwd())
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.call(["open", str(path)])
            else:
                subprocess.call(["xdg-open", str(path)])
        except Exception as exc:  # pragma: no cover - UI safeguard
            messagebox.showerror("Error", f"Could not open folder:\n{exc}")

    # Layout
    top_frame = tk.Frame(root)
    top_frame.pack(fill="x", padx=10, pady=10)

    tk.Label(top_frame, text="Free-text Feature Description:").pack(anchor="w")
    description_text = scrolledtext.ScrolledText(root, height=10, wrap=tk.WORD)
    description_text.pack(fill="both", expand=False, padx=10)

    controls_frame = tk.Frame(root)
    controls_frame.pack(fill="x", padx=10, pady=8)

    tk.Label(controls_frame, text="Output folder:").grid(row=0, column=0, sticky="w")
    tk.Entry(controls_frame, textvariable=output_dir, width=50).grid(
        row=0, column=1, sticky="we", padx=(5, 5)
    )
    tk.Button(controls_frame, text="Browse…", command=choose_output_dir).grid(
        row=0, column=2, padx=(0, 5)
    )
    tk.Label(controls_frame, text="OpenAI API Key (optional):").grid(row=1, column=0, sticky="w")
    tk.Entry(controls_frame, textvariable=api_key_var, show="*", width=50).grid(
        row=1, column=1, sticky="we", padx=(5, 5)
    )
    tk.Button(controls_frame, text="Generate Excel & Preview", command=generate).grid(
        row=0, column=3, rowspan=2, padx=(5, 0)
    )
    tk.Button(controls_frame, text="Open Folder", command=open_output_folder).grid(
        row=0, column=4, rowspan=2, padx=(5, 0)
    )
    controls_frame.columnconfigure(1, weight=1)

    tk.Label(root, text="Textual Preview (for verification before opening Excel):").pack(
        anchor="w", padx=10
    )
    preview_text = scrolledtext.ScrolledText(root, height=18, wrap=tk.WORD, state="disabled")
    preview_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    status_var = tk.StringVar(value="Ready.")
    status_bar = tk.Label(root, textvariable=status_var, anchor="w")
    status_bar.pack(fill="x", side="bottom")

    root.mainloop()


if __name__ == "__main__":
    # Default to GUI; allow CLI mode via: python feature_sizing_app.py --cli
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        main()
    else:
        run_gui()

