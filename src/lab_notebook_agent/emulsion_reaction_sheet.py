from __future__ import annotations

from copy import copy
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


SOURCE_SPREADSHEET_ID = "1DmLIodsRi1sMhaHrB8jHRv4LU-2HknKD4vKm7b8_29c"
SOURCE_SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    f"{SOURCE_SPREADSHEET_ID}/edit"
)
SOURCE_SHEET = "CCSP-52"
MODEL_VERSION = "ccsp-emulsion-reaction.v1"


@dataclass(frozen=True)
class ReactionCharge:
    stage: str
    addition: str
    material: str
    role: str
    source_cell: str
    phr: float | None = None
    stage_factor_g_per_phr: float | None = None
    allocation_fraction: float = 1.0
    direct_mass_g: float | None = None
    density_g_ml: float | None = None
    active_fraction: float = 1.0
    molecular_weight_g_mol: float | None = None
    polymer_forming: bool = False
    nonvolatile: bool = False
    carry_from_stage: str | None = None
    note: str = ""


@dataclass(frozen=True)
class CalculatedCharge:
    charge: ReactionCharge
    mass_g: float
    volume_ml: float
    active_mass_g: float
    polymer_forming_mass_g: float
    nonvolatile_mass_g: float
    active_moles: float | None


STAGE_ORDER = ("Seed", "Core", "Shell", "Functional Shell")


def _scaled(
    stage: str,
    addition: str,
    material: str,
    role: str,
    source_cell: str,
    phr: float,
    factor: float,
    density: float,
    *,
    allocation: float = 1.0,
    active: float = 1.0,
    mw: float | None = None,
    polymer: bool = False,
    nonvolatile: bool = False,
    note: str = "",
) -> ReactionCharge:
    return ReactionCharge(
        stage=stage,
        addition=addition,
        material=material,
        role=role,
        source_cell=source_cell,
        phr=phr,
        stage_factor_g_per_phr=factor,
        allocation_fraction=allocation,
        density_g_ml=density,
        active_fraction=active,
        molecular_weight_g_mol=mw,
        polymer_forming=polymer,
        nonvolatile=nonvolatile,
        note=note,
    )


def _carry(
    stage: str,
    addition: str,
    material: str,
    source_cell: str,
    prior_stage: str,
    *,
    direct_mass_g: float | None = None,
    density: float | None = None,
    note: str = "",
) -> ReactionCharge:
    return ReactionCharge(
        stage=stage,
        addition=addition,
        material=material,
        role="latex carry",
        source_cell=source_cell,
        direct_mass_g=direct_mass_g,
        density_g_ml=density,
        carry_from_stage=prior_stage,
        note=note,
    )


def ccsp_52_charges() -> tuple[ReactionCharge, ...]:
    """Return the non-zero CCSP-52 planning charges with corrected identities.

    The source workbook contains many zero placeholders and three large time-feed
    tables. This first deterministic contract keeps only charges that change the
    batch, while retaining source-cell provenance for every value.
    """

    seed = 0.6
    shell = 0.084
    functional = 0.168
    return (
        _scaled("Seed", "Pre-reactor", "DI Water", "water", "D3", 180, seed, 1.0, allocation=0.85),
        _scaled("Seed", "Pre-reactor", "Butyl Acrylate", "monomer", "D4", 99.7, seed, 0.89, allocation=0.10, mw=128.17, polymer=True, nonvolatile=True),
        _scaled("Seed", "Pre-reactor", "Dowfax 2A1", "surfactant", "D7", 1, seed, 1.0, nonvolatile=True, note="Source treats solution as 100% active; verify product actives."),
        _scaled("Seed", "Monomer pre-emulsion", "Butyl Acrylate", "monomer", "D8", 99.7, seed, 0.89, allocation=0.90, mw=128.17, polymer=True, nonvolatile=True),
        _scaled("Seed", "Monomer pre-emulsion", "1,4-Butanediol Dimethacrylate", "crosslinker", "D9", 0.1, seed, 1.051, mw=198.26, polymer=True, nonvolatile=True),
        _scaled("Seed", "Monomer pre-emulsion", "Allyl Methacrylate", "crosslinker", "D10", 0.2, seed, 0.938, mw=126.15, polymer=True, nonvolatile=True),
        _scaled("Seed", "Aqueous pre-emulsion", "Dowfax 2A1", "surfactant", "D11", 2, seed, 1.0, nonvolatile=True, note="Source treats solution as 100% active; verify product actives."),
        _scaled("Seed", "Aqueous pre-emulsion", "DI Water", "water", "D12", 300, seed, 1.0, allocation=0.15),
        _scaled("Seed", "Redox shot", "APS solution", "initiator", "D13", 1, seed, 1.0, active=0.052, mw=228.18, nonvolatile=True),
        _scaled("Seed", "Redox shot", "FF6 solution", "reductant", "D14", 1, seed, 1.0, active=0.03, nonvolatile=True),
        _scaled("Seed", "Redox feed", "APS solution", "initiator", "D15", 9, seed, 1.0, active=0.052, mw=228.18, nonvolatile=True),
        _scaled("Seed", "Redox feed", "FF6 solution", "reductant", "D16", 9, seed, 1.0, active=0.03, nonvolatile=True),
        _scaled("Seed", "Chase", "TBHP solution", "initiator", "D17", 3.3, seed, 1.0, active=0.048, mw=90.12, nonvolatile=True),
        _scaled("Seed", "Chase", "FF6 solution", "reductant", "D18", 3.3, seed, 1.0, active=0.048, nonvolatile=True),

        _carry("Core", "Pre-reactor", "Seed Latex", "E37", "Seed", direct_mass_g=11.4, density=1.0, note="Aliquot from the seed batch; mass is an explicit source input."),
        _scaled("Core", "Pre-reactor", "DI Water", "water", "D38", 50, seed, 1.0),
        _scaled("Core", "Monomer pre-emulsion", "Butyl Acrylate", "monomer", "D43", 99.7, seed, 0.89, mw=128.17, polymer=True, nonvolatile=True),
        _scaled("Core", "Monomer pre-emulsion", "Allyl Methacrylate", "crosslinker", "D45", 0.2, seed, 0.938, mw=126.15, polymer=True, nonvolatile=True),
        _scaled("Core", "Monomer pre-emulsion", "1,4-Butanediol Dimethacrylate", "crosslinker", "D46", 0.1, seed, 1.051, mw=198.26, polymer=True, nonvolatile=True),
        _scaled("Core", "Monomer pre-emulsion", "tert-Dodecyl Mercaptan", "chain-transfer agent", "D47", 0.03, seed, 0.85, mw=202.40, nonvolatile=True),
        _scaled("Core", "Monomer pre-emulsion", "Aerosol MA-80", "surfactant", "D48", 2, seed, 1.0, active=0.80, nonvolatile=True),
        _scaled("Core", "Aqueous pre-emulsion", "DI Water", "water", "D50", 50, seed, 1.0),
        _scaled("Core", "Redox feed", "APS solution", "initiator", "D54", 10, seed, 1.0, active=0.026, mw=228.18, nonvolatile=True, note="Correct main initiator for core CTA:I and M:I."),
        _scaled("Core", "Redox feed", "FF6 solution", "reductant", "D55", 10, seed, 1.0, active=0.015, nonvolatile=True),

        _carry("Shell", "Pre-reactor", "Core Latex", "E71", "Core"),
        _scaled("Shell", "Monomer pre-emulsion", "Methyl Methacrylate", "monomer", "D74", 87.3, shell, 0.94, mw=100.12, polymer=True, nonvolatile=True, note="Density normalized to 0.94 g/mL; source F74 uses 0.89 while F101 uses 0.94."),
        _scaled("Shell", "Monomer pre-emulsion", "Allyl Methacrylate", "crosslinker", "D75", 0.5, shell, 0.938, mw=126.15, polymer=True, nonvolatile=True),
        _scaled("Shell", "Monomer pre-emulsion", "1,4-Butanediol Dimethacrylate", "crosslinker", "D76", 0.2, shell, 1.051, mw=198.26, polymer=True, nonvolatile=True),
        _scaled("Shell", "Monomer pre-emulsion", "Ethyl Acrylate", "monomer", "D77", 12, shell, 0.94, mw=100.12, polymer=True, nonvolatile=True),
        _scaled("Shell", "Monomer pre-emulsion", "tert-Dodecyl Mercaptan", "chain-transfer agent", "D78", 0.1, shell, 0.85, mw=202.40, nonvolatile=True),
        _scaled("Shell", "Monomer pre-emulsion", "Aerosol MA-80", "surfactant", "D79", 1, shell, 1.0, active=0.80, nonvolatile=True),
        _scaled("Shell", "Aqueous pre-emulsion", "Surfactant A solution", "surfactant", "D80", 0.1, shell, 1.0, active=0.30, nonvolatile=True),
        _scaled("Shell", "Aqueous pre-emulsion", "DI Water", "water", "D81", 50, shell, 1.0),
        _scaled("Shell", "Redox feed", "TBHP solution", "initiator", "D85", 30, shell, 1.0, active=0.009, mw=90.12, nonvolatile=True),
        _scaled("Shell", "Redox feed", "FF6 solution", "reductant", "D86", 30, shell, 1.0, active=0.005, nonvolatile=True),

        _carry("Functional Shell", "Pre-reactor", "Core-Shell Latex", "E97", "Shell"),
        _scaled("Functional Shell", "Pre-reactor", "Sodium Bicarbonate solution", "buffer", "D98", 8, functional, 1.0, active=0.025, nonvolatile=True),
        _scaled("Functional Shell", "Monomer pre-emulsion", "Methyl Methacrylate", "monomer", "D101", 58, functional, 0.94, mw=100.12, polymer=True, nonvolatile=True),
        _scaled("Functional Shell", "Monomer pre-emulsion", "Ethyl Acrylate", "monomer", "D102", 4, functional, 0.94, mw=100.12, polymer=True, nonvolatile=True),
        _scaled("Functional Shell", "Monomer pre-emulsion", "Glycidyl Methacrylate", "functional monomer", "D103", 38, functional, 1.48, mw=142.15, polymer=True, nonvolatile=True),
        _scaled("Functional Shell", "Monomer pre-emulsion", "Aerosol MA-80", "surfactant", "D105", 1, functional, 1.0, active=0.80, nonvolatile=True),
        _scaled("Functional Shell", "Aqueous pre-emulsion", "Surfactant A solution", "surfactant", "D106", 0.1, functional, 1.0, active=0.30, nonvolatile=True),
        _scaled("Functional Shell", "Aqueous pre-emulsion", "DI Water", "water", "D107", 50, functional, 1.0),
        _scaled("Functional Shell", "Aqueous pre-emulsion", "Sipomer COPS solution", "polymerizable surfactant", "D109", 1, functional, 1.0, active=0.20, polymer=True, nonvolatile=True),
        _scaled("Functional Shell", "Aqueous pre-emulsion", "Rhodapex CL 910 solution", "polymerizable surfactant", "D110", 0.5, functional, 1.0, active=0.20, polymer=True, nonvolatile=True),
        _scaled("Functional Shell", "Redox feed", "APS solution", "initiator", "D111", 25, functional, 1.0, active=0.009, mw=228.18, nonvolatile=True),
        _scaled("Functional Shell", "Redox feed", "SFS solution", "reductant", "D112", 25, functional, 1.0, active=0.005, nonvolatile=True),
        _scaled("Functional Shell", "Chase", "TBHP solution", "initiator", "D113", 25, shell, 1.0, active=0.009, mw=90.12, nonvolatile=True, note="Source deliberately uses the 0.084 chase factor, not the 0.168 functional-shell factor."),
        _scaled("Functional Shell", "Chase", "FF6 solution", "reductant", "D114", 25, shell, 1.0, active=0.005, nonvolatile=True, note="Source deliberately uses the 0.084 chase factor, not the 0.168 functional-shell factor."),
    )


def calculate_reaction_plan(
    charges: tuple[ReactionCharge, ...] | None = None,
) -> dict[str, Any]:
    charges = charges or ccsp_52_charges()
    calculated: list[CalculatedCharge] = []
    stages: dict[str, dict[str, float]] = {}
    failures: list[str] = []
    warnings: list[str] = []

    for stage in STAGE_ORDER:
        stage_rows = [row for row in charges if row.stage == stage]
        for row in stage_rows:
            if row.carry_from_stage:
                prior = stages.get(row.carry_from_stage)
                if prior is None:
                    raise ValueError(f"Carry stage {row.carry_from_stage!r} has not been calculated")
                if row.direct_mass_g is None:
                    mass = prior["total_mass_g"]
                    volume = prior["total_volume_ml"]
                    polymer_mass = prior["polymer_forming_mass_g"]
                    nonvolatile_mass = prior["nonvolatile_mass_g"]
                    active_mass = nonvolatile_mass
                else:
                    mass = row.direct_mass_g
                    volume = mass / row.density_g_ml if row.density_g_ml else mass
                    fraction = mass / prior["total_mass_g"] if prior["total_mass_g"] else 0.0
                    polymer_mass = prior["polymer_forming_mass_g"] * fraction
                    nonvolatile_mass = prior["nonvolatile_mass_g"] * fraction
                    active_mass = nonvolatile_mass
            elif row.direct_mass_g is not None:
                mass = row.direct_mass_g
                if not row.density_g_ml or row.density_g_ml <= 0:
                    failures.append(f"{stage}/{row.material}: direct charge requires a positive density")
                    volume = 0.0
                else:
                    volume = mass / row.density_g_ml
                active_mass = mass * row.active_fraction
                polymer_mass = active_mass if row.polymer_forming else 0.0
                nonvolatile_mass = active_mass if row.nonvolatile else 0.0
            else:
                if row.phr is None or row.stage_factor_g_per_phr is None:
                    failures.append(f"{stage}/{row.material}: missing PHR or stage factor")
                    mass = 0.0
                else:
                    mass = row.phr * row.stage_factor_g_per_phr * row.allocation_fraction
                if not row.density_g_ml or row.density_g_ml <= 0:
                    failures.append(f"{stage}/{row.material}: density must be positive")
                    volume = 0.0
                else:
                    volume = mass / row.density_g_ml
                if not 0 <= row.active_fraction <= 1:
                    failures.append(f"{stage}/{row.material}: active fraction must be between 0 and 1")
                active_mass = mass * row.active_fraction
                polymer_mass = active_mass if row.polymer_forming else 0.0
                nonvolatile_mass = active_mass if row.nonvolatile else 0.0

            active_moles = None
            if row.molecular_weight_g_mol:
                active_moles = active_mass / row.molecular_weight_g_mol
            calculated.append(
                CalculatedCharge(
                    charge=row,
                    mass_g=mass,
                    volume_ml=volume,
                    active_mass_g=active_mass,
                    polymer_forming_mass_g=polymer_mass,
                    nonvolatile_mass_g=nonvolatile_mass,
                    active_moles=active_moles,
                )
            )

        rows = [row for row in calculated if row.charge.stage == stage]
        total_mass = sum(row.mass_g for row in rows)
        total_volume = sum(row.volume_ml for row in rows)
        polymer_mass = sum(row.polymer_forming_mass_g for row in rows)
        nonvolatile_mass = sum(row.nonvolatile_mass_g for row in rows)
        stages[stage] = {
            "total_mass_g": total_mass,
            "total_volume_ml": total_volume,
            "polymer_forming_mass_g": polymer_mass,
            "nonvolatile_mass_g": nonvolatile_mass,
            "polymer_solids_fraction": polymer_mass / total_mass if total_mass else 0.0,
            "nonvolatile_solids_fraction": nonvolatile_mass / total_mass if total_mass else 0.0,
        }

    core_rows = [row for row in calculated if row.charge.stage == "Core"]
    core_monomer_moles = sum(
        row.active_moles or 0.0
        for row in core_rows
        if row.charge.polymer_forming
    )
    crosslinker_moles = sum(
        row.active_moles or 0.0
        for row in core_rows
        if row.charge.role == "crosslinker"
    )
    cta_moles = sum(
        row.active_moles or 0.0
        for row in core_rows
        if row.charge.role == "chain-transfer agent"
    )
    main_initiator_moles = sum(
        row.active_moles or 0.0
        for row in core_rows
        if row.charge.material == "APS solution" and row.charge.addition == "Redox feed"
    )
    if main_initiator_moles <= 0:
        failures.append("Core main initiator moles must be positive")

    new_polymer = {}
    for stage in ("Core", "Shell", "Functional Shell"):
        new_polymer[stage] = sum(
            row.polymer_forming_mass_g
            for row in calculated
            if row.charge.stage == stage and not row.charge.carry_from_stage
        )
    new_polymer_total = sum(new_polymer.values())
    actual_stage_fractions = {
        stage: mass / new_polymer_total if new_polymer_total else 0.0
        for stage, mass in new_polymer.items()
    }
    target_stage_fractions = {"Core": 0.70, "Shell": 0.10, "Functional Shell": 0.20}
    maximum_fraction_deviation = max(
        abs(actual_stage_fractions[stage] - target_stage_fractions[stage])
        for stage in target_stage_fractions
    )
    if maximum_fraction_deviation > 0.005:
        warnings.append("Actual core/shell/functional-shell fractions differ from targets by more than 0.5 percentage point")
    if any("verify product actives" in row.charge.note.lower() for row in calculated):
        warnings.append("Dowfax 2A1 active fraction is not controlled in the source workbook")

    ratios = {
        "core_monomer_moles": core_monomer_moles,
        "core_main_initiator_moles": main_initiator_moles,
        "core_cta_moles": cta_moles,
        "core_crosslinker_moles": crosslinker_moles,
        "core_monomer_to_initiator": core_monomer_moles / main_initiator_moles if main_initiator_moles else None,
        "core_cta_to_initiator": cta_moles / main_initiator_moles if main_initiator_moles else None,
        "core_crosslinker_to_cta": crosslinker_moles / cta_moles if cta_moles else None,
    }

    return {
        "schema": MODEL_VERSION,
        "source": {
            "spreadsheet_id": SOURCE_SPREADSHEET_ID,
            "spreadsheet_url": SOURCE_SPREADSHEET_URL,
            "sheet": SOURCE_SHEET,
        },
        "charges": [
            {
                **asdict(row.charge),
                "mass_g": row.mass_g,
                "volume_ml": row.volume_ml,
                "active_mass_g": row.active_mass_g,
                "polymer_forming_mass_g": row.polymer_forming_mass_g,
                "nonvolatile_mass_g": row.nonvolatile_mass_g,
                "active_moles": row.active_moles,
            }
            for row in calculated
        ],
        "stages": stages,
        "new_polymer_mass_g": new_polymer,
        "actual_stage_fractions": actual_stage_fractions,
        "target_stage_fractions": target_stage_fractions,
        "ratios": ratios,
        "checks": {
            "ready": not failures,
            "failures": failures,
            "warnings": warnings,
            "maximum_stage_fraction_deviation": maximum_fraction_deviation,
        },
    }


def ccsp_source_audit() -> list[dict[str, str]]:
    return [
        {
            "severity": "error",
            "source_cell": "CCSP-52!G6",
            "finding": "Zero-mass sodium acetate reports 0.1277 mL because the formula is =E10/F10.",
            "deterministic_rule": "Every ordinary charge volume is its own mass divided by its own density; this row would be 0 mL.",
        },
        {
            "severity": "error",
            "source_cell": "CCSP-52!K35:K36",
            "finding": "CTA:I and M:I are #DIV/0! because they reference zero KPS/chase cells and omit the non-zero APS main feed at E54.",
            "deterministic_rule": "Use active APS mass from E54 times 2.6 wt%, divided by APS MW 228.18 g/mol.",
        },
        {
            "severity": "error",
            "source_cell": "CCSP-52!K36",
            "finding": "The historical M:I formula divides initiator active mass by 92 while CTA:I uses 270; neither matches the APS identity used in CCSP-52.",
            "deterministic_rule": "One controlled initiator identity and molecular weight feeds every molar ratio.",
        },
        {
            "severity": "warning",
            "source_cell": "CCSP-52!F74",
            "finding": "Methyl methacrylate density is 0.89 g/mL in the shell but 0.94 g/mL in the functional shell.",
            "deterministic_rule": "Normalize the material to one controlled density; the plan uses 0.94 g/mL pending lot/SDS confirmation.",
        },
        {
            "severity": "warning",
            "source_cell": "CCSP-52!K46,K54",
            "finding": "Rounded shell factors 0.084 and 0.168 yield about 70.4/9.9/19.7 instead of the labeled 70/10/20 split.",
            "deterministic_rule": "Show target and actual stage fractions separately; do not label rounded actuals as exact targets.",
        },
        {
            "severity": "warning",
            "source_cell": "CCSP-52!K48,K56",
            "finding": "Source solids formulas omit at least the shell BDDMA charge and mix polymer mass with nonvolatile-solids meaning.",
            "deterministic_rule": "Report polymer-forming mass fraction and theoretical nonvolatile-solids fraction as separate outputs.",
        },
        {
            "severity": "warning",
            "source_cell": "CCSP-52!D7,D11",
            "finding": "Dowfax 2A1 is treated as 100% active without a controlled solution concentration.",
            "deterministic_rule": "Keep the source treatment visible but require product-actives confirmation before release.",
        },
    ]


def build_ccsp_audit_report() -> dict[str, Any]:
    plan = calculate_reaction_plan()
    return {
        "schema": "ccsp-emulsion-reaction-audit.v1",
        "source": plan["source"],
        "important_values": [
            "run identity and operator",
            "core/shell/functional-shell target fractions",
            "stage, addition group, material identity, PHR, scale factor, and allocation",
            "planned mass, controlled density, and calculated volume",
            "solution active fraction, molecular weight, active mass, and active moles",
            "stage total mass and volume",
            "polymer-forming and theoretical nonvolatile-solids fractions",
            "main initiator, CTA, crosslinker, and monomer molar ratios",
            "reaction temperature, feed duration, and target/result particle size",
        ],
        "excluded_from_first_sheet": [
            "zero-quantity placeholder materials",
            "instantaneous radical-flux model",
            "syringe-diameter lookup table",
            "three repeated time-feed matrices",
            "advanced surfactant CMC mixing model",
        ],
        "source_findings": ccsp_source_audit(),
        "source_comparison": [
            {
                "quantity": "Core total mass",
                "source_cell": "CCSP-52!C66",
                "source_value": 144.618,
                "deterministic_value": plan["stages"]["Core"]["total_mass_g"],
                "status": "matches",
            },
            {
                "quantity": "Shell total mass",
                "source_cell": "CCSP-52!C92",
                "source_value": 162.3588,
                "deterministic_value": plan["stages"]["Shell"]["total_mass_g"],
                "status": "matches",
            },
            {
                "quantity": "Functional-shell final mass",
                "source_cell": "CCSP-52!C118",
                "source_value": 201.9396,
                "deterministic_value": plan["stages"]["Functional Shell"]["total_mass_g"],
                "status": "matches",
            },
            {
                "quantity": "Functional-shell final volume",
                "source_cell": "CCSP-52!C119",
                "source_value": 208.90969864221154,
                "deterministic_value": plan["stages"]["Functional Shell"]["total_volume_ml"],
                "status": "corrected MMA density",
            },
            {
                "quantity": "Core CTA:I",
                "source_cell": "CCSP-52!K35",
                "source_value": "#DIV/0!",
                "deterministic_value": plan["ratios"]["core_cta_to_initiator"],
                "status": "corrected initiator reference and MW",
            },
            {
                "quantity": "Core M:I",
                "source_cell": "CCSP-52!K36",
                "source_value": "#DIV/0!",
                "deterministic_value": plan["ratios"]["core_monomer_to_initiator"],
                "status": "corrected initiator reference and MW",
            },
        ],
        "deterministic_result": {
            "stages": plan["stages"],
            "target_stage_fractions": plan["target_stage_fractions"],
            "actual_stage_fractions": plan["actual_stage_fractions"],
            "ratios": plan["ratios"],
            "checks": plan["checks"],
        },
    }


def _save_ccsp_reaction_workbook_dense(output: str | Path) -> Path:
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    plan = calculate_reaction_plan()
    audit = ccsp_source_audit()
    charges = ccsp_52_charges()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Reaction Sheet"
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A17"

    navy = "17365D"
    blue = "D9EAF7"
    input_fill = "FFF2CC"
    calc_fill = "E2F0D9"
    warning_fill = "FCE4D6"
    sheet["A1"] = "CCSP EMULSION POLYMERIZATION — DETERMINISTIC REACTION PLAN"
    sheet["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    sheet["A1"].fill = PatternFill("solid", fgColor=navy)
    sheet.merge_cells("A1:P1")
    sheet["A2"] = "First-pass planning sheet: formulation quantities, unit-safe calculations, stage totals, ratios, and explicit checks."
    sheet.merge_cells("A2:P2")
    sheet["A2"].fill = PatternFill("solid", fgColor=blue)

    summary_headers = ["Stage", "Total mass (g)", "Total volume (mL)", "Polymer-forming mass (g)", "Nonvolatile mass (g)", "Polymer solids (%)", "Nonvolatile solids (%)"]
    for column, value in enumerate(summary_headers, start=1):
        cell = sheet.cell(4, column, value)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(bold=True, color="FFFFFF")
    for row_number, stage in enumerate(STAGE_ORDER, start=5):
        values = plan["stages"][stage]
        sheet.cell(row_number, 1, stage)
        sheet.cell(row_number, 2, values["total_mass_g"])
        sheet.cell(row_number, 3, values["total_volume_ml"])
        sheet.cell(row_number, 4, values["polymer_forming_mass_g"])
        sheet.cell(row_number, 5, values["nonvolatile_mass_g"])
        sheet.cell(row_number, 6, values["polymer_solids_fraction"])
        sheet.cell(row_number, 7, values["nonvolatile_solids_fraction"])
        for column in range(2, 8):
            sheet.cell(row_number, column).fill = PatternFill("solid", fgColor=calc_fill)
    for row in range(5, 9):
        for column in (2, 3, 4, 5):
            sheet.cell(row, column).number_format = "0.000"
        for column in (6, 7):
            sheet.cell(row, column).number_format = "0.00%"

    sheet["I4"] = "Core molar checks"
    sheet["I4"].fill = PatternFill("solid", fgColor=navy)
    sheet["I4"].font = Font(bold=True, color="FFFFFF")
    ratio_rows = (
        ("Monomer:initiator", "core_monomer_to_initiator"),
        ("CTA:initiator", "core_cta_to_initiator"),
        ("Crosslinker:CTA", "core_crosslinker_to_cta"),
    )
    for row_number, (label, key) in enumerate(ratio_rows, start=5):
        sheet.cell(row_number, 9, label)
        sheet.cell(row_number, 10, plan["ratios"][key])
        sheet.cell(row_number, 10).number_format = "0.000"
        sheet.cell(row_number, 10).fill = PatternFill("solid", fgColor=calc_fill)
    sheet["L4"] = "Design split"
    sheet["L4"].fill = PatternFill("solid", fgColor=navy)
    sheet["L4"].font = Font(bold=True, color="FFFFFF")
    for row_number, stage in enumerate(("Core", "Shell", "Functional Shell"), start=5):
        sheet.cell(row_number, 12, stage)
        sheet.cell(row_number, 13, plan["target_stage_fractions"][stage])
        sheet.cell(row_number, 14, plan["actual_stage_fractions"][stage])
        sheet.cell(row_number, 13).number_format = "0.00%"
        sheet.cell(row_number, 14).number_format = "0.00%"
    sheet["M4"] = "Target"
    sheet["N4"] = "Actual"
    sheet["P4"] = "Source"
    sheet["P5"] = SOURCE_SHEET
    sheet["P6"] = SOURCE_SPREADSHEET_URL
    sheet["P6"].hyperlink = SOURCE_SPREADSHEET_URL

    sheet["A10"] = "Release checks"
    sheet["A10"].fill = PatternFill("solid", fgColor=navy)
    sheet["A10"].font = Font(bold=True, color="FFFFFF")
    check_rows = [
        ("Deterministic calculation", "PASS" if plan["checks"]["ready"] else "FAIL"),
        ("Design split within 0.5 percentage point", "PASS" if plan["checks"]["maximum_stage_fraction_deviation"] <= 0.005 else "WARN"),
        ("Dowfax product actives confirmed", "VERIFY"),
        ("Core initiator identity and MW", "APS / 228.18 g/mol"),
    ]
    for row_number, (label, value) in enumerate(check_rows, start=11):
        sheet.cell(row_number, 1, label)
        sheet.cell(row_number, 2, value)
        sheet.cell(row_number, 2).fill = PatternFill("solid", fgColor=calc_fill if value == "PASS" else warning_fill)

    table_headers = (
        "Stage", "Addition", "Material", "Role", "PHR", "Stage factor", "Allocation", "Direct/carry mass (g)",
        "Planned mass (g)", "Density (g/mL)", "Volume (mL)", "Active fraction", "Active mass (g)",
        "MW (g/mol)", "Active moles", "Source / note",
    )
    header_row = 16
    for column, value in enumerate(table_headers, start=1):
        cell = sheet.cell(header_row, column, value)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    stage_summary_rows = {stage: row for row, stage in enumerate(STAGE_ORDER, start=5)}
    table_start = header_row + 1
    for index, charge in enumerate(charges, start=table_start):
        sheet.cell(index, 1, charge.stage)
        sheet.cell(index, 2, charge.addition)
        sheet.cell(index, 3, charge.material)
        sheet.cell(index, 4, charge.role)
        sheet.cell(index, 5, charge.phr)
        sheet.cell(index, 6, charge.stage_factor_g_per_phr)
        sheet.cell(index, 7, charge.allocation_fraction)
        if charge.carry_from_stage and charge.direct_mass_g is None:
            sheet.cell(index, 8, f"=$B${stage_summary_rows[charge.carry_from_stage]}")
        elif charge.direct_mass_g is not None:
            sheet.cell(index, 8, charge.direct_mass_g)
        sheet.cell(index, 9, f'=IF(H{index}<>"",H{index},E{index}*F{index}*G{index})')
        sheet.cell(index, 10, charge.density_g_ml)
        if charge.carry_from_stage and charge.direct_mass_g is None:
            sheet.cell(index, 11, f"=$C${stage_summary_rows[charge.carry_from_stage]}")
        else:
            sheet.cell(index, 11, f'=IFERROR(I{index}/J{index},"")')
        if charge.carry_from_stage:
            prior_row = stage_summary_rows[charge.carry_from_stage]
            sheet.cell(index, 12, f'=IFERROR($E${prior_row}/$B${prior_row},0)')
        else:
            sheet.cell(index, 12, charge.active_fraction if charge.nonvolatile else 0.0)
        sheet.cell(index, 13, f"=I{index}*L{index}")
        sheet.cell(index, 14, charge.molecular_weight_g_mol)
        sheet.cell(index, 15, f'=IFERROR(M{index}/N{index},"")')
        source_note = f"{SOURCE_SHEET}!{charge.source_cell}"
        if charge.note:
            source_note += f" — {charge.note}"
        sheet.cell(index, 16, source_note)
        for column in (5, 6, 7, 8, 10, 12, 14):
            sheet.cell(index, column).fill = PatternFill("solid", fgColor=input_fill)
        for column in (9, 11, 13, 15):
            sheet.cell(index, column).fill = PatternFill("solid", fgColor=calc_fill)
        for column in range(1, 17):
            sheet.cell(index, column).alignment = Alignment(vertical="top", wrap_text=True)

    table_end = table_start + len(charges) - 1

    polymer_roles = '{"monomer","functional monomer","crosslinker","polymerizable surfactant"}'
    stage_new_polymer_formulas: dict[str, str] = {}
    for stage in ("Core", "Shell", "Functional Shell"):
        stage_new_polymer_formulas[stage] = (
            f'SUM(SUMIFS($M${table_start}:$M${table_end},$A${table_start}:$A${table_end},'
            f'"{stage}",$D${table_start}:$D${table_end},{polymer_roles}))'
        )
    design_denominator = "+".join(stage_new_polymer_formulas.values())

    for row_number, stage in enumerate(STAGE_ORDER, start=5):
        sheet.cell(row_number, 2, f'=SUMIF($A${table_start}:$A${table_end},A{row_number},$I${table_start}:$I${table_end})')
        sheet.cell(row_number, 3, f'=SUMIF($A${table_start}:$A${table_end},A{row_number},$K${table_start}:$K${table_end})')
        new_polymer_formula = (
            f'SUM(SUMIFS($M${table_start}:$M${table_end},$A${table_start}:$A${table_end},'
            f'A{row_number},$D${table_start}:$D${table_end},{polymer_roles}))'
        )
        if stage == "Seed":
            polymer_formula = new_polymer_formula
        else:
            prior_stage = STAGE_ORDER[STAGE_ORDER.index(stage) - 1]
            prior_row = stage_summary_rows[prior_stage]
            polymer_formula = (
                f'{new_polymer_formula}+SUMIFS($I${table_start}:$I${table_end},'
                f'$A${table_start}:$A${table_end},A{row_number},$D${table_start}:$D${table_end},'
                f'"latex carry")*$D${prior_row}/$B${prior_row}'
            )
        sheet.cell(row_number, 4, f"={polymer_formula}")
        sheet.cell(row_number, 5, f'=SUMIF($A${table_start}:$A${table_end},A{row_number},$M${table_start}:$M${table_end})')
        sheet.cell(row_number, 6, f'=IFERROR(D{row_number}/B{row_number},0)')
        sheet.cell(row_number, 7, f'=IFERROR(E{row_number}/B{row_number},0)')

    sheet["J5"] = (
        f'=IFERROR(SUM(SUMIFS($O${table_start}:$O${table_end},$A${table_start}:$A${table_end},"Core",'
        f'$D${table_start}:$D${table_end},{polymer_roles}))/SUMIFS($O${table_start}:$O${table_end},'
        f'$A${table_start}:$A${table_end},"Core",$C${table_start}:$C${table_end},"APS solution",'
        f'$B${table_start}:$B${table_end},"Redox feed"),"")'
    )
    sheet["J6"] = (
        f'=IFERROR(SUMIFS($O${table_start}:$O${table_end},$A${table_start}:$A${table_end},"Core",'
        f'$D${table_start}:$D${table_end},"chain-transfer agent")/SUMIFS($O${table_start}:$O${table_end},'
        f'$A${table_start}:$A${table_end},"Core",$C${table_start}:$C${table_end},"APS solution",'
        f'$B${table_start}:$B${table_end},"Redox feed"),"")'
    )
    sheet["J7"] = (
        f'=IFERROR(SUMIFS($O${table_start}:$O${table_end},$A${table_start}:$A${table_end},"Core",'
        f'$D${table_start}:$D${table_end},"crosslinker")/SUMIFS($O${table_start}:$O${table_end},'
        f'$A${table_start}:$A${table_end},"Core",$D${table_start}:$D${table_end},'
        f'"chain-transfer agent"),"")'
    )
    for row_number, stage in enumerate(("Core", "Shell", "Functional Shell"), start=5):
        sheet.cell(row_number, 14, f'=IFERROR(({stage_new_polymer_formulas[stage]})/({design_denominator}),0)')
    sheet["B12"] = '=IF(MAX(ABS(M5-N5),ABS(M6-N6),ABS(M7-N7))<=0.005,"PASS","WARN")'

    for column in (5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15):
        for row in range(table_start, table_end + 1):
            sheet.cell(row, column).number_format = "0.000000"
    sheet.auto_filter.ref = f"A{header_row}:P{table_end}"

    audit_start = table_end + 3
    sheet.cell(audit_start, 1, "SOURCE CALCULATION AUDIT")
    sheet.cell(audit_start, 1).fill = PatternFill("solid", fgColor=navy)
    sheet.cell(audit_start, 1).font = Font(bold=True, color="FFFFFF")
    for column, value in enumerate(("Severity", "Source cell", "Finding", "Deterministic rule"), start=1):
        cell = sheet.cell(audit_start + 1, column, value)
        cell.fill = PatternFill("solid", fgColor=blue)
        cell.font = Font(bold=True)
    for row_number, finding in enumerate(audit, start=audit_start + 2):
        for column, key in enumerate(("severity", "source_cell", "finding", "deterministic_rule"), start=1):
            sheet.cell(row_number, column, finding[key])
            sheet.cell(row_number, column).alignment = Alignment(wrap_text=True, vertical="top")

    widths = {
        "A": 18, "B": 23, "C": 32, "D": 22, "E": 10, "F": 12, "G": 11, "H": 18,
        "I": 18, "J": 15, "K": 15, "L": 14, "M": 16, "N": 14, "O": 14, "P": 50,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    sheet.conditional_formatting.add(
        f"B11:B14",
        CellIsRule(operator="equal", formula=['"FAIL"'], fill=PatternFill("solid", fgColor="F4CCCC")),
    )
    sheet.print_title_rows = "$1:$16"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    workbook.save(output)
    return output


def _title(sheet: Any, text: str, subtitle: str, end_column: int) -> None:
    navy = "17365D"
    pale = "EAF2F8"
    sheet.sheet_view.showGridLines = False
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_column)
    sheet["A1"] = text
    sheet["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    sheet["A1"].fill = PatternFill("solid", fgColor=navy)
    sheet["A1"].alignment = Alignment(vertical="center")
    sheet.row_dimensions[1].height = 30
    sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_column)
    sheet["A2"] = subtitle
    sheet["A2"].fill = PatternFill("solid", fgColor=pale)
    sheet["A2"].font = Font(color="44546A")
    sheet["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[2].height = 28


def _section_header(sheet: Any, row: int, start: int, end: int, text: str) -> None:
    sheet.merge_cells(start_row=row, start_column=start, end_row=row, end_column=end)
    cell = sheet.cell(row, start, text)
    cell.fill = PatternFill("solid", fgColor="2F75B5")
    cell.font = Font(bold=True, color="FFFFFF")
    cell.alignment = Alignment(vertical="center")


def _table_header(sheet: Any, row: int, headers: tuple[str, ...]) -> None:
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(row, column, header)
        cell.fill = PatternFill("solid", fgColor="5B9BD5")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[row].height = 30


def save_ccsp_reaction_workbook(output: str | Path) -> Path:
    """Save the visually separated deterministic CCSP planning workbook."""

    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    charges = ccsp_52_charges()
    plan = calculate_reaction_plan(charges)
    audit = ccsp_source_audit()
    workbook = Workbook()
    workbook.remove(workbook.active)

    start = workbook.create_sheet("Start Here")
    recipe = workbook.create_sheet("Recipe")
    summary = workbook.create_sheet("Stage Summary")
    chemistry = workbook.create_sheet("Chemistry")
    assumptions = workbook.create_sheet("Assumptions")

    input_fill = PatternFill("solid", fgColor="FFF2CC")
    calc_fill = PatternFill("solid", fgColor="E2F0D9")
    warning_fill = PatternFill("solid", fgColor="FCE4D6")
    good_fill = PatternFill("solid", fgColor="C6E0B4")
    neutral_fill = PatternFill("solid", fgColor="D9EAF7")
    stage_colors = {
        "Seed": "DDEBF7",
        "Core": "E2F0D9",
        "Shell": "FFF2CC",
        "Functional Shell": "FCE4D6",
    }

    # Recipe: the primary working surface.
    _title(
        recipe,
        "CCSP RECIPE",
        "Edit yellow cells only. Mass and volume recalculate from PHR × scale × allocation, or from an explicit carry mass.",
        10,
    )
    recipe_headers = (
        "Stage", "Addition", "Material", "PHR", "Scale", "Allocation", "Carry / direct mass (g)",
        "Mass (g)", "Density (g/mL)", "Volume (mL)",
    )
    _table_header(recipe, 4, recipe_headers)
    recipe_start = 5
    recipe_rows: dict[int, ReactionCharge] = {}
    stage_first_rows: dict[str, int] = {}
    for row_number, charge in enumerate(charges, start=recipe_start):
        recipe_rows[row_number] = charge
        stage_first_rows.setdefault(charge.stage, row_number)
        recipe.cell(row_number, 1, charge.stage)
        recipe.cell(row_number, 2, charge.addition)
        recipe.cell(row_number, 3, charge.material)
        recipe.cell(row_number, 4, charge.phr)
        recipe.cell(row_number, 5, charge.stage_factor_g_per_phr)
        recipe.cell(row_number, 6, charge.allocation_fraction)
        if charge.carry_from_stage and charge.direct_mass_g is None:
            prior_summary_row = STAGE_ORDER.index(charge.carry_from_stage) + 5
            recipe.cell(row_number, 7, f"='Stage Summary'!$B${prior_summary_row}")
        elif charge.direct_mass_g is not None:
            recipe.cell(row_number, 7, charge.direct_mass_g)
        recipe.cell(row_number, 8, f'=IF(G{row_number}<>"",G{row_number},D{row_number}*E{row_number}*F{row_number})')
        recipe.cell(row_number, 9, charge.density_g_ml)
        if charge.carry_from_stage and charge.direct_mass_g is None:
            prior_summary_row = STAGE_ORDER.index(charge.carry_from_stage) + 5
            recipe.cell(row_number, 10, f"='Stage Summary'!$C${prior_summary_row}")
        else:
            recipe.cell(row_number, 10, f'=IFERROR(H{row_number}/I{row_number},"")')
        for column in (4, 5, 6, 7, 9):
            recipe.cell(row_number, column).fill = input_fill
        for column in (8, 10):
            recipe.cell(row_number, column).fill = calc_fill
        for column in range(1, 11):
            recipe.cell(row_number, column).alignment = Alignment(vertical="center")
        recipe.cell(row_number, 1).fill = PatternFill("solid", fgColor=stage_colors[charge.stage])
    recipe_end = recipe_start + len(charges) - 1
    recipe_table = Table(displayName="CCSPRecipe", ref=f"A4:J{recipe_end}")
    recipe_table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    recipe.add_table(recipe_table)
    recipe.freeze_panes = "D5"
    for row in range(recipe_start, recipe_end + 1):
        for column in (4, 5, 7, 8, 9, 10):
            recipe.cell(row, column).number_format = "0.000"
        recipe.cell(row, 6).number_format = "0.0%"
    for column, width in {
        "A": 19, "B": 22, "C": 32, "D": 10, "E": 10, "F": 11, "G": 18, "H": 13, "I": 15, "J": 14,
    }.items():
        recipe.column_dimensions[column].width = width
    for stage, first_row in stage_first_rows.items():
        recipe.row_dimensions[first_row].height = 25

    # Chemistry: properties and mole calculations kept out of the recipe view.
    _title(
        chemistry,
        "CHEMISTRY DETAILS",
        "Active fractions, molecular weights, and source provenance. Yellow values are assumptions or controlled material properties.",
        10,
    )
    chemistry_headers = (
        "Stage", "Material", "Role", "Active / nonvolatile fraction", "Active mass (g)", "MW (g/mol)",
        "Active moles", "Polymer-forming mass (g)", "Source cell", "Short note",
    )
    _table_header(chemistry, 4, chemistry_headers)
    polymer_roles = {"monomer", "functional monomer", "crosslinker", "polymerizable surfactant"}
    for row_number, charge in recipe_rows.items():
        chemistry.cell(row_number, 1, f"=Recipe!A{row_number}")
        chemistry.cell(row_number, 2, f"=Recipe!C{row_number}")
        chemistry.cell(row_number, 3, charge.role)
        if charge.carry_from_stage:
            prior_summary_row = STAGE_ORDER.index(charge.carry_from_stage) + 5
            chemistry.cell(row_number, 4, f"='Stage Summary'!$G${prior_summary_row}")
        else:
            chemistry.cell(row_number, 4, charge.active_fraction if charge.nonvolatile else 0.0)
        chemistry.cell(row_number, 5, f"=Recipe!H{row_number}*D{row_number}")
        chemistry.cell(row_number, 6, charge.molecular_weight_g_mol)
        chemistry.cell(row_number, 7, f'=IFERROR(E{row_number}/F{row_number},"")')
        if charge.carry_from_stage:
            prior_summary_row = STAGE_ORDER.index(charge.carry_from_stage) + 5
            chemistry.cell(row_number, 8, f"=Recipe!H{row_number}*'Stage Summary'!$F${prior_summary_row}")
        elif charge.role in polymer_roles:
            chemistry.cell(row_number, 8, f"=E{row_number}")
        else:
            chemistry.cell(row_number, 8, 0.0)
        chemistry.cell(row_number, 9, f"{SOURCE_SHEET}!{charge.source_cell}")
        short_note = charge.note.split(";")[0] if charge.note else ""
        chemistry.cell(row_number, 10, short_note)
        for column in (4, 6):
            chemistry.cell(row_number, column).fill = input_fill
        for column in (5, 7, 8):
            chemistry.cell(row_number, column).fill = calc_fill
        for column in range(1, 11):
            chemistry.cell(row_number, column).alignment = Alignment(vertical="top", wrap_text=column == 10)
    chemistry_table = Table(displayName="CCSPChemistry", ref=f"A4:J{recipe_end}")
    chemistry_table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium4", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    chemistry.add_table(chemistry_table)
    chemistry.freeze_panes = "D5"
    for row in range(recipe_start, recipe_end + 1):
        chemistry.cell(row, 4).number_format = "0.0%"
        for column in (5, 6, 7, 8):
            chemistry.cell(row, column).number_format = "0.000000"
    for column, width in {
        "A": 19, "B": 31, "C": 24, "D": 22, "E": 16, "F": 14, "G": 15, "H": 22, "I": 16, "J": 48,
    }.items():
        chemistry.column_dimensions[column].width = width

    # Stage Summary: concise calculation dashboard.
    _title(
        summary,
        "STAGE SUMMARY",
        "Totals and design checks. Polymer solids and nonvolatile solids are deliberately separate definitions.",
        9,
    )
    summary_headers = (
        "Stage", "Mass (g)", "Volume (mL)", "Polymer mass (g)", "Nonvolatile mass (g)",
        "Polymer solids", "Nonvolatile solids", "Target split", "Actual split",
    )
    _table_header(summary, 4, summary_headers)
    for row_number, stage in enumerate(STAGE_ORDER, start=5):
        summary.cell(row_number, 1, stage)
        summary.cell(row_number, 2, f'=SUMIF(Recipe!$A$5:$A${recipe_end},A{row_number},Recipe!$H$5:$H${recipe_end})')
        summary.cell(row_number, 3, f'=SUMIF(Recipe!$A$5:$A${recipe_end},A{row_number},Recipe!$J$5:$J${recipe_end})')
        summary.cell(row_number, 4, f'=SUMIF(Chemistry!$A$5:$A${recipe_end},A{row_number},Chemistry!$H$5:$H${recipe_end})')
        summary.cell(row_number, 5, f'=SUMIF(Chemistry!$A$5:$A${recipe_end},A{row_number},Chemistry!$E$5:$E${recipe_end})')
        summary.cell(row_number, 6, f'=IFERROR(D{row_number}/B{row_number},0)')
        summary.cell(row_number, 7, f'=IFERROR(E{row_number}/B{row_number},0)')
        if stage == "Seed":
            summary.cell(row_number, 8, "—")
            summary.cell(row_number, 9, "—")
        else:
            summary.cell(row_number, 8, plan["target_stage_fractions"][stage])
            summary.cell(
                row_number,
                9,
                f'=IFERROR(SUMIFS(Chemistry!$H$5:$H${recipe_end},Chemistry!$A$5:$A${recipe_end},A{row_number},'
                f'Chemistry!$C$5:$C${recipe_end},"<>latex carry")/'
                f'SUMIFS(Chemistry!$H$5:$H${recipe_end},Chemistry!$A$5:$A${recipe_end},"<>Seed",'
                f'Chemistry!$C$5:$C${recipe_end},"<>latex carry"),0)',
            )
        summary.cell(row_number, 1).fill = PatternFill("solid", fgColor=stage_colors[stage])
        for column in range(2, 10):
            summary.cell(row_number, column).fill = calc_fill
    for row in range(5, 9):
        for column in (2, 3, 4, 5):
            summary.cell(row, column).number_format = "0.000"
        for column in (6, 7, 8, 9):
            summary.cell(row, column).number_format = "0.00%"
    _section_header(summary, 11, 1, 3, "Core molar ratios")
    monomer_moles = (
        f'SUMIFS(Chemistry!$G$5:$G${recipe_end},Chemistry!$A$5:$A${recipe_end},"Core",Chemistry!$C$5:$C${recipe_end},"monomer")+'
        f'SUMIFS(Chemistry!$G$5:$G${recipe_end},Chemistry!$A$5:$A${recipe_end},"Core",Chemistry!$C$5:$C${recipe_end},"functional monomer")+'
        f'SUMIFS(Chemistry!$G$5:$G${recipe_end},Chemistry!$A$5:$A${recipe_end},"Core",Chemistry!$C$5:$C${recipe_end},"crosslinker")+'
        f'SUMIFS(Chemistry!$G$5:$G${recipe_end},Chemistry!$A$5:$A${recipe_end},"Core",Chemistry!$C$5:$C${recipe_end},"polymerizable surfactant")'
    )
    initiator_moles = f'SUMIFS(Chemistry!$G$5:$G${recipe_end},Chemistry!$A$5:$A${recipe_end},"Core",Chemistry!$B$5:$B${recipe_end},"APS solution")'
    cta_moles = f'SUMIFS(Chemistry!$G$5:$G${recipe_end},Chemistry!$A$5:$A${recipe_end},"Core",Chemistry!$C$5:$C${recipe_end},"chain-transfer agent")'
    crosslinker_moles = f'SUMIFS(Chemistry!$G$5:$G${recipe_end},Chemistry!$A$5:$A${recipe_end},"Core",Chemistry!$C$5:$C${recipe_end},"crosslinker")'
    ratio_rows = (
        (12, "Monomer : initiator", f"=IFERROR(({monomer_moles})/({initiator_moles}),0)"),
        (13, "CTA : initiator", f"=IFERROR(({cta_moles})/({initiator_moles}),0)"),
        (14, "Crosslinker : CTA", f"=IFERROR(({crosslinker_moles})/({cta_moles}),0)"),
    )
    for row_number, label, formula in ratio_rows:
        summary.cell(row_number, 1, label)
        summary.cell(row_number, 2, formula)
        summary.cell(row_number, 2).number_format = "0.000"
        summary.cell(row_number, 2).fill = calc_fill
    summary["E11"] = "Design check"
    summary["E11"].fill = PatternFill("solid", fgColor="2F75B5")
    summary["E11"].font = Font(bold=True, color="FFFFFF")
    summary["E12"] = "Max target deviation"
    summary["F12"] = "=MAX(ABS(I6-H6),ABS(I7-H7),ABS(I8-H8))"
    summary["F12"].number_format = "0.00%"
    summary["E13"] = "Calculation status"
    summary["F13"] = '=IF(F12<=0.005,"PASS","WARN")'
    summary["E14"] = "Assumption status"
    summary["F14"] = "OPEN"
    summary["F13"].fill = good_fill
    summary["F14"].fill = warning_fill
    summary.freeze_panes = "A5"
    for column, width in {"A": 23, "B": 15, "C": 15, "D": 18, "E": 20, "F": 18, "G": 19, "H": 14, "I": 14}.items():
        summary.column_dimensions[column].width = width

    # Start Here: one-page decision surface with no implementation prose.
    _title(
        start,
        "CCSP EMULSION POLYMERIZATION",
        "Use this page to orient. Enter or revise quantities on Recipe; confirm chemistry properties on Chemistry.",
        8,
    )
    _section_header(start, 4, 1, 4, "1 · Plan the batch")
    start["A5"] = "Edit"
    start["B5"] = "Recipe"
    start["B5"].hyperlink = "#'Recipe'!A1"
    start["A6"] = "Check properties"
    start["B6"] = "Chemistry"
    start["B6"].hyperlink = "#'Chemistry'!A1"
    start["A7"] = "Review totals"
    start["B7"] = "Stage Summary"
    start["B7"].hyperlink = "#'Stage Summary'!A1"
    start["A8"] = "Resolve"
    start["B8"] = "Assumptions"
    start["B8"].hyperlink = "#'Assumptions'!A1"
    _section_header(start, 4, 5, 8, "2 · Release status")
    status_rows = (
        (5, "Calculations", "='Stage Summary'!F13", good_fill),
        (6, "Design split", "='Stage Summary'!F13", good_fill),
        (7, "Material assumptions", "OPEN", warning_fill),
        (8, "Ready for controlled use", "NO", warning_fill),
    )
    for row_number, label, value, fill in status_rows:
        start.cell(row_number, 5, label)
        start.cell(row_number, 6, value)
        start.cell(row_number, 6).fill = fill
        start.cell(row_number, 6).font = Font(bold=True)
    _section_header(start, 11, 1, 8, "Batch at a glance")
    glance_headers = ("Stage", "Mass", "Volume", "Polymer solids", "Nonvolatile solids", "Target", "Actual", "Difference")
    _table_header(start, 12, glance_headers)
    for row_number, stage in enumerate(STAGE_ORDER, start=13):
        source_row = STAGE_ORDER.index(stage) + 5
        start.cell(row_number, 1, stage)
        start.cell(row_number, 2, f"='Stage Summary'!B{source_row}")
        start.cell(row_number, 3, f"='Stage Summary'!C{source_row}")
        start.cell(row_number, 4, f"='Stage Summary'!F{source_row}")
        start.cell(row_number, 5, f"='Stage Summary'!G{source_row}")
        start.cell(row_number, 6, f"='Stage Summary'!H{source_row}")
        start.cell(row_number, 7, f"='Stage Summary'!I{source_row}")
        start.cell(row_number, 8, f'=IFERROR(G{row_number}-F{row_number},"")')
        start.cell(row_number, 1).fill = PatternFill("solid", fgColor=stage_colors[stage])
    for row in range(13, 17):
        for column in (2, 3):
            start.cell(row, column).number_format = "0.000"
        for column in (4, 5, 6, 7, 8):
            start.cell(row, column).number_format = "0.00%"
    _section_header(start, 19, 1, 4, "Core ratios")
    for row_number, (label, source_row) in enumerate(
        (("Monomer : initiator", 12), ("CTA : initiator", 13), ("Crosslinker : CTA", 14)),
        start=20,
    ):
        start.cell(row_number, 1, label)
        start.cell(row_number, 2, f"='Stage Summary'!B{source_row}")
        start.cell(row_number, 2).number_format = "0.000"
        start.cell(row_number, 2).fill = calc_fill
    _section_header(start, 19, 5, 8, "Open confirmations")
    start["E20"] = "Dowfax actives"
    start["F20"] = "VERIFY"
    start["E21"] = "MMA density"
    start["F21"] = "VERIFY 0.94 g/mL"
    start["E22"] = "Final controlled review"
    start["F22"] = "OPEN"
    for row in range(20, 23):
        start.cell(row, 6).fill = warning_fill
    start["A25"] = "Source"
    start["B25"] = f"{SOURCE_SHEET} · live shared workbook"
    start["B25"].hyperlink = SOURCE_SPREADSHEET_URL
    start.freeze_panes = "A4"
    start.page_setup.fitToWidth = 1
    start.sheet_properties.pageSetUpPr.fitToPage = True
    for column, width in {"A": 22, "B": 22, "C": 16, "D": 18, "E": 25, "F": 18, "G": 14, "H": 14}.items():
        start.column_dimensions[column].width = width

    # Assumptions: concise issue register; full prose remains in the audit JSON/docs.
    _title(
        assumptions,
        "ASSUMPTIONS & SOURCE ISSUES",
        "Resolve VERIFY items before treating the plan as controlled. Full evidence is in the companion audit report.",
        6,
    )
    assumption_headers = ("Status", "Topic", "Source", "Problem", "Decision", "Owner / evidence")
    _table_header(assumptions, 4, assumption_headers)
    concise = (
        ("FIXED", "Row-local volume", "G6", "Wrong-row volume reference", "Mass ÷ same-row density", "Calculation test"),
        ("FIXED", "Core initiator ratios", "K35:K36", "Zero denominator / wrong MW", "APS at E54; MW 228.18", "Calculation test"),
        ("VERIFY", "MMA density", "F74 / F101", "0.89 vs 0.94 g/mL", "Use 0.94 provisionally", "Lot SDS"),
        ("VISIBLE", "Stage split", "K46 / K54", "Rounded factors shift split", "Show target and actual", "Stage Summary"),
        ("FIXED", "Solids definition", "K48 / K56", "Ambiguous and incomplete", "Show polymer and nonvolatile", "Calculation test"),
        ("VERIFY", "Dowfax actives", "D7 / D11", "No controlled active fraction", "Source assumption retained", "Product TDS / lot COA"),
    )
    for row_number, values in enumerate(concise, start=5):
        for column, value in enumerate(values, start=1):
            assumptions.cell(row_number, column, value)
            assumptions.cell(row_number, column).alignment = Alignment(wrap_text=True, vertical="top")
        assumptions.cell(row_number, 1).fill = good_fill if values[0] in {"FIXED", "VISIBLE"} else warning_fill
        assumptions.cell(row_number, 1).font = Font(bold=True)
    assumptions["A13"] = "Full source audit"
    assumptions["B13"] = "See ccsp_emulsion_reaction_audit_v1.json and docs/ccsp-emulsion-reaction-sheet-v1.md"
    assumptions.merge_cells("B13:F13")
    assumptions.freeze_panes = "A5"
    for column, width in {"A": 13, "B": 23, "C": 16, "D": 31, "E": 32, "F": 25}.items():
        assumptions.column_dimensions[column].width = width

    for sheet in workbook.worksheets:
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_margins.left = 0.25
        sheet.page_margins.right = 0.25
        sheet.page_margins.top = 0.4
        sheet.page_margins.bottom = 0.4
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None and cell.font == Font():
                    cell.font = Font(name="Arial", size=10)

    workbook.active = 0
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    workbook.save(output)
    return output


def _feed_schedule_rows() -> tuple[tuple[str, float, float, float, float], ...]:
    """Return the CCSP-52 interval schedule in Vivek's compact feed-table shape."""

    return (
        ("Core", 15, 0.10, 0.06, 0.00),
        ("Core", 30, 0.10, 0.00, 0.06),
        ("Core", 45, 0.30, 0.15, 0.15),
        ("Core", 60, 0.50, 0.15, 0.15),
        ("Core", 90, 0.70, 0.16, 0.16),
        ("Core", 120, 0.90, 0.16, 0.16),
        ("Core", 150, 1.00, 0.16, 0.16),
        ("Core", 180, 1.00, 0.16, 0.16),
        ("Core", 210, 1.00, 0.00, 0.00),
        ("Core", 240, 1.00, 0.00, 0.00),
        ("Core", 270, 1.00, 0.00, 0.00),
        ("Shell", 15, 0.40, 0.06, 0.00),
        ("Shell", 30, 0.40, 0.00, 0.06),
        ("Shell", 45, 0.60, 0.23, 0.23),
        ("Shell", 60, 0.80, 0.23, 0.23),
        ("Shell", 75, 1.00, 0.24, 0.24),
        ("Shell", 90, 1.00, 0.24, 0.24),
        ("Shell", 120, 1.00, 0.00, 0.00),
        ("Shell", 150, 1.00, 0.00, 0.00),
        ("Functional Shell", 15, 0.40, 0.06, 0.00),
        ("Functional Shell", 30, 0.40, 0.00, 0.06),
        ("Functional Shell", 45, 0.60, 0.23, 0.23),
        ("Functional Shell", 60, 0.80, 0.23, 0.23),
        ("Functional Shell", 75, 1.00, 0.24, 0.24),
        ("Functional Shell", 90, 1.00, 0.24, 0.24),
        ("Functional Shell", 120, 1.00, 0.00, 0.00),
        ("Functional Shell", 150, 1.00, 0.00, 0.00),
    )


def save_ccsp_reaction_workbook(output: str | Path) -> Path:
    """Save a simple, color-coded reaction plan with an explicit feed schedule."""

    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    charges = ccsp_52_charges()
    workbook = Workbook()
    reaction = workbook.active
    reaction.title = "Reaction Plan"
    feeds = workbook.create_sheet("Feed Schedule")
    checks = workbook.create_sheet("Checks")
    assumptions = workbook.create_sheet("Assumptions")

    navy = "1F4E78"
    header_blue = "5B9BD5"
    input_yellow = PatternFill("solid", fgColor="FFF2CC")
    output_green = PatternFill("solid", fgColor="E2F0D9")
    warning_orange = PatternFill("solid", fgColor="FCE4D6")
    stage_colors = {
        "Seed": "DDEBF7",
        "Core": "E2F0D9",
        "Shell": "FFF2CC",
        "Functional Shell": "FCE4D6",
    }

    def title(sheet: Any, text: str, subtitle: str, end_column: int) -> None:
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_column)
        sheet["A1"] = text
        sheet["A1"].fill = PatternFill("solid", fgColor=navy)
        sheet["A1"].font = Font(name="Arial", size=18, bold=True, color="FFFFFF")
        sheet["A1"].alignment = Alignment(vertical="center")
        sheet.row_dimensions[1].height = 30
        sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_column)
        sheet["A2"] = subtitle
        sheet["A2"].font = Font(name="Arial", size=10, italic=True, color="44546A")
        sheet["A2"].alignment = Alignment(wrap_text=True, vertical="center")
        sheet.row_dimensions[2].height = 30

    def header(sheet: Any, row: int, values: tuple[str, ...]) -> None:
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row, column, value)
            cell.fill = PatternFill("solid", fgColor=header_blue)
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        sheet.row_dimensions[row].height = 32

    # Reaction Plan: one primary bench-facing table.
    title(
        reaction,
        "CCSP REACTION PLAN",
        "Yellow = input. Green = calculated output. Edit the recipe here; use Feed Schedule to plan pumps and timed additions.",
        10,
    )
    reaction["A3"] = "INPUT"
    reaction["A3"].fill = input_yellow
    reaction["B3"] = "CALCULATED"
    reaction["B3"].fill = output_green
    reaction["D3"] = "Source"
    reaction["E3"] = "CCSP-52 quantities; layout informed by all 219 Vivek EP tabs"
    reaction["E3"].hyperlink = "https://docs.google.com/spreadsheets/d/1zhdi6LkRNhQRitZjzd1mkN4kOrtrMKiQ2Wx4H_Hws0c/edit"
    headers = (
        "Stage", "Addition", "Material", "PHR", "Factor", "Allocation", "Direct / carry mass (g)",
        "Density (g/mL)", "Mass (g)", "Volume (mL)",
    )
    header(reaction, 5, headers)
    first_rows: dict[str, int] = {}
    recipe_end = 5 + len(charges)
    for row_number, charge in enumerate(charges, start=6):
        first_rows.setdefault(charge.stage, row_number)
        reaction.cell(row_number, 1, charge.stage)
        reaction.cell(row_number, 2, charge.addition)
        reaction.cell(row_number, 3, charge.material)
        reaction.cell(row_number, 4, charge.phr)
        reaction.cell(row_number, 5, charge.stage_factor_g_per_phr)
        reaction.cell(row_number, 6, charge.allocation_fraction)
        if charge.carry_from_stage and charge.direct_mass_g is None:
            prior = STAGE_ORDER.index(charge.carry_from_stage) + 6
            reaction.cell(row_number, 7, f"=Checks!$B${prior}")
        elif charge.direct_mass_g is not None:
            reaction.cell(row_number, 7, charge.direct_mass_g)
        reaction.cell(row_number, 8, charge.density_g_ml)
        reaction.cell(row_number, 9, f'=IF(G{row_number}<>"",G{row_number},D{row_number}*E{row_number}*F{row_number})')
        if charge.carry_from_stage and charge.direct_mass_g is None:
            prior = STAGE_ORDER.index(charge.carry_from_stage) + 6
            reaction.cell(row_number, 10, f"=Checks!$C${prior}")
        else:
            reaction.cell(row_number, 10, f'=IFERROR(I{row_number}/H{row_number},"")')
        for column in (4, 5, 6, 7, 8):
            reaction.cell(row_number, column).fill = input_yellow
        for column in (9, 10):
            reaction.cell(row_number, column).fill = output_green
        reaction.cell(row_number, 1).fill = PatternFill("solid", fgColor=stage_colors[charge.stage])
        for column in range(1, 11):
            reaction.cell(row_number, column).alignment = Alignment(vertical="center")
    reaction.freeze_panes = "D6"
    reaction.auto_filter.ref = f"A5:J{recipe_end}"
    for row in range(6, recipe_end + 1):
        for column in (4, 5, 7, 8, 9, 10):
            reaction.cell(row, column).number_format = "0.000"
        reaction.cell(row, 6).number_format = "0.0%"
    for column, width in {"A": 19, "B": 23, "C": 34, "D": 10, "E": 10, "F": 11, "G": 20, "H": 16, "I": 13, "J": 14}.items():
        reaction.column_dimensions[column].width = width

    # Checks is deliberately small; detailed chemistry stays in the JSON audit.
    title(checks, "STAGE CHECKS", "Totals and key planning checks only.", 7)
    header(checks, 5, ("Stage", "Mass (g)", "Volume (mL)", "Polymer mass (g)", "Polymer solids", "Target split", "Actual split"))
    polymer_roles = {"monomer", "functional monomer", "crosslinker", "polymerizable surfactant"}
    for row_number, stage in enumerate(STAGE_ORDER, start=6):
        checks.cell(row_number, 1, stage)
        checks.cell(row_number, 2, f'=SUMIF(\'Reaction Plan\'!$A$6:$A${recipe_end},A{row_number},\'Reaction Plan\'!$I$6:$I${recipe_end})')
        checks.cell(row_number, 3, f'=SUMIF(\'Reaction Plan\'!$A$6:$A${recipe_end},A{row_number},\'Reaction Plan\'!$J$6:$J${recipe_end})')
        role_rows = [i for i, charge in enumerate(charges, start=6) if charge.stage == stage and charge.role in polymer_roles]
        checks.cell(row_number, 4, "=" + "+".join(f"'Reaction Plan'!I{i}*{charges[i-6].active_fraction}" for i in role_rows) if role_rows else 0)
        carry_row = first_rows[stage]
        checks.cell(row_number, 5, f'=IFERROR((D{row_number}+IF(A{row_number}="Seed",0,\'Reaction Plan\'!G{carry_row}*INDEX($E$6:$E$9,MATCH(A{row_number},$A$6:$A$9,0)-1)))/B{row_number},0)')
        if stage == "Seed":
            checks.cell(row_number, 6, "—")
            checks.cell(row_number, 7, "—")
        else:
            checks.cell(row_number, 6, {"Core": 0.70, "Shell": 0.10, "Functional Shell": 0.20}[stage])
            checks.cell(row_number, 7, f'=IFERROR(D{row_number}/SUM($D$7:$D$9),0)')
        checks.cell(row_number, 1).fill = PatternFill("solid", fgColor=stage_colors[stage])
        for column in range(2, 8):
            checks.cell(row_number, column).fill = output_green
    checks["A12"] = "Max split deviation"
    checks["B12"] = "=MAX(ABS(G7-F7),ABS(G8-F8),ABS(G9-F9))"
    checks["A13"] = "Calculation status"
    checks["B13"] = '=IF(B12<=0.005,"PASS","CHECK")'
    checks["B13"].fill = output_green
    for row in range(6, 10):
        for column in (2, 3, 4):
            checks.cell(row, column).number_format = "0.000"
        for column in (5, 6, 7):
            checks.cell(row, column).number_format = "0.00%"
    checks["B12"].number_format = "0.00%"
    for column, width in {"A": 22, "B": 15, "C": 16, "D": 18, "E": 16, "F": 14, "G": 14}.items():
        checks.column_dimensions[column].width = width
    checks.freeze_panes = "A6"

    # Feed Schedule: the missing operational layer, modeled on Vivek's time tables.
    title(
        feeds,
        "FEED SCHEDULE",
        "Yellow values define timing and distribution. Green cells calculate the pump rates from the current Reaction Plan.",
        10,
    )
    feeds["A3"] = "INPUT"
    feeds["A3"].fill = input_yellow
    feeds["B3"] = "CALCULATED"
    feeds["B3"].fill = output_green
    feed_headers = (
        "Stage", "End time (min)", "Interval (min)", "Emulsion cumulative %", "Oxidant interval %",
        "Reductant interval %", "Emulsion added (mL)", "Emulsion rate (mL/min)",
        "Oxidant rate (mL/min)", "Reductant rate (mL/min)",
    )
    header(feeds, 5, feed_headers)
    schedule = _feed_schedule_rows()
    for row_number, (stage, end_min, emulsion_cumulative, oxidant_fraction, reductant_fraction) in enumerate(schedule, start=6):
        feeds.cell(row_number, 1, stage)
        feeds.cell(row_number, 2, end_min)
        previous_row = row_number - 1
        if row_number == 6 or schedule[row_number - 7][0] != stage:
            feeds.cell(row_number, 3, f"=B{row_number}")
            prior_cumulative = "0"
        else:
            feeds.cell(row_number, 3, f"=B{row_number}-B{previous_row}")
            prior_cumulative = f"D{previous_row}"
        feeds.cell(row_number, 4, emulsion_cumulative)
        feeds.cell(row_number, 5, oxidant_fraction)
        feeds.cell(row_number, 6, reductant_fraction)
        total_emulsion = (
            f'SUMIFS(\'Reaction Plan\'!$J$6:$J${recipe_end},\'Reaction Plan\'!$A$6:$A${recipe_end},A{row_number},'
            f'\'Reaction Plan\'!$B$6:$B${recipe_end},"Monomer pre-emulsion")+'
            f'SUMIFS(\'Reaction Plan\'!$J$6:$J${recipe_end},\'Reaction Plan\'!$A$6:$A${recipe_end},A{row_number},'
            f'\'Reaction Plan\'!$B$6:$B${recipe_end},"Aqueous pre-emulsion")'
        )
        oxidant_material = 'IF(A{r}="Shell","TBHP solution","APS solution")'.format(r=row_number)
        reductant_material = 'IF(A{r}="Functional Shell","SFS solution","FF6 solution")'.format(r=row_number)
        feeds.cell(row_number, 7, f'=({total_emulsion})*(D{row_number}-{prior_cumulative})')
        feeds.cell(row_number, 8, f'=IFERROR(G{row_number}/C{row_number},0)')
        feeds.cell(
            row_number,
            9,
            f'=IFERROR(E{row_number}*SUMIFS(\'Reaction Plan\'!$J$6:$J${recipe_end},\'Reaction Plan\'!$A$6:$A${recipe_end},A{row_number},'
            f'\'Reaction Plan\'!$B$6:$B${recipe_end},"Redox feed",\'Reaction Plan\'!$C$6:$C${recipe_end},{oxidant_material})/C{row_number},0)',
        )
        feeds.cell(
            row_number,
            10,
            f'=IFERROR(F{row_number}*SUMIFS(\'Reaction Plan\'!$J$6:$J${recipe_end},\'Reaction Plan\'!$A$6:$A${recipe_end},A{row_number},'
            f'\'Reaction Plan\'!$B$6:$B${recipe_end},"Redox feed",\'Reaction Plan\'!$C$6:$C${recipe_end},{reductant_material})/C{row_number},0)',
        )
        for column in (2, 4, 5, 6):
            feeds.cell(row_number, column).fill = input_yellow
        for column in (3, 7, 8, 9, 10):
            feeds.cell(row_number, column).fill = output_green
        feeds.cell(row_number, 1).fill = PatternFill("solid", fgColor=stage_colors[stage])
    feed_end = 5 + len(schedule)
    feeds.auto_filter.ref = f"A5:J{feed_end}"
    feeds.freeze_panes = "D6"
    for row in range(6, feed_end + 1):
        for column in (4, 5, 6):
            feeds.cell(row, column).number_format = "0.0%"
        for column in (7, 8, 9, 10):
            feeds.cell(row, column).number_format = "0.000"
    for column, width in {"A": 19, "B": 14, "C": 14, "D": 21, "E": 19, "F": 20, "G": 19, "H": 20, "I": 20, "J": 21}.items():
        feeds.column_dimensions[column].width = width

    title(assumptions, "ASSUMPTIONS", "Only unresolved decisions and source corrections are shown here.", 5)
    header(assumptions, 5, ("Status", "Topic", "Source", "Working decision", "Needed evidence"))
    rows = (
        ("VERIFY", "MMA density", "CCSP-52 F74/F101", "Use 0.94 g/mL", "Current lot SDS"),
        ("VERIFY", "Dowfax actives", "CCSP-52 D7/D11", "Retain source basis", "Product TDS or lot COA"),
        ("VISIBLE", "Stage split", "CCSP-52 factors", "Show target and actual", "Checks tab"),
        ("FIXED", "Volume formulas", "CCSP-52 G6", "Mass divided by same-row density", "Automated test"),
        ("FIXED", "Initiator ratios", "CCSP-52 K35:K36", "Use nonzero main APS feed", "Automated test"),
        ("SOURCE", "Feed schedule", "CCSP-52 rows 122:153", "Editable interval schedule", "Operator review"),
    )
    for row_number, values in enumerate(rows, start=6):
        for column, value in enumerate(values, start=1):
            assumptions.cell(row_number, column, value)
            assumptions.cell(row_number, column).alignment = Alignment(wrap_text=True, vertical="top")
        assumptions.cell(row_number, 1).fill = warning_orange if values[0] == "VERIFY" else output_green
    for column, width in {"A": 13, "B": 25, "C": 24, "D": 35, "E": 28}.items():
        assumptions.column_dimensions[column].width = width
    assumptions.freeze_panes = "A6"

    for sheet in workbook.worksheets:
        sheet.sheet_view.showGridLines = False
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_margins.left = 0.25
        sheet.page_margins.right = 0.25
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    updated_font = copy(cell.font)
                    updated_font.name = "Arial"
                    cell.font = updated_font

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    workbook.save(output)
    return output
