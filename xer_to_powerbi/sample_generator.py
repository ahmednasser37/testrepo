"""
Generates a realistic Primavera P6 .XER sample file for
Al-Narjis Water Network - Phase 2, Riyadh (NWC standards).
"""

import random
from datetime import datetime, timedelta
from pathlib import Path


# ── Configuration ────────────────────────────────────────────────────────────
PROJECT_START = datetime(2024, 10, 1)
PROJECT_END   = datetime(2026, 4, 1)
DATA_DATE     = datetime(2026, 4, 1)

ZONES = [f"Zone {i:02d}" for i in range(1, 9)]

DISCIPLINES = [
    "Excavation",
    "Pipe Laying",
    "Backfill",
    "Reinstatement",
    "Testing & Commissioning",
]

RESOURCES = [
    ("RSRC-001", "HDPE Pipe 315mm",        "MT", "m",   185.0),
    ("RSRC-002", "HDPE Pipe 250mm",        "MT", "m",   140.0),
    ("RSRC-003", "HDPE Pipe 160mm",        "MT", "m",    85.0),
    ("RSRC-004", "Concrete C25",           "MT", "m3",  420.0),
    ("RSRC-005", "Sand Bedding",           "MT", "m3",   28.0),
    ("RSRC-006", "Mechanical Excavation",  "MT", "m3",   35.0),
    ("RSRC-007", "Asphalt Reinstatement",  "MT", "m2",   95.0),
    ("RSRC-008", "Gate Valve DN300",       "MT", "EA",  4200.0),
]

# Zone completion profiles: (planned_pct, actual_pct)
ZONE_PROFILES = {
    "Zone 01": (1.00, 1.00),   # complete
    "Zone 02": (1.00, 0.95),   # almost done
    "Zone 03": (0.90, 0.80),   # in progress, slightly behind
    "Zone 04": (0.75, 0.60),   # in progress, behind
    "Zone 05": (0.65, 0.50),   # in progress, behind
    "Zone 06": (0.40, 0.25),   # early stages, behind
    "Zone 07": (0.15, 0.05),   # just started
    "Zone 08": (0.00, 0.00),   # not started
}

DISC_SEQUENCE = {
    "Excavation":            0,
    "Pipe Laying":           1,
    "Backfill":              2,
    "Reinstatement":         3,
    "Testing & Commissioning": 4,
}

# Which resources belong to which discipline
DISC_RESOURCES = {
    "Excavation":              ["RSRC-006"],
    "Pipe Laying":             ["RSRC-001", "RSRC-002", "RSRC-003", "RSRC-008"],
    "Backfill":                ["RSRC-005", "RSRC-004"],
    "Reinstatement":           ["RSRC-007"],
    "Testing & Commissioning": [],
}


def _fmt(dt: datetime | None) -> str:
    if dt is None:
        return ""
    return dt.strftime("%Y-%m-%d %H:%M")


def _add_days(dt: datetime, days: float) -> datetime:
    return dt + timedelta(days=days)


# ── Row builders ─────────────────────────────────────────────────────────────

def _project_row() -> dict:
    return {
        "proj_id":         "PROJ-001",
        "proj_short_name": "ALNARJIS-P2",
        "proj_name":       "Al-Narjis Water Network - Phase 2",
        "plan_start_date": _fmt(PROJECT_START),
        "plan_end_date":   _fmt(PROJECT_END),
        "data_date":       _fmt(DATA_DATE),
        "last_recalc_date":_fmt(DATA_DATE),
    }


def _wbs_rows() -> list[dict]:
    rows = []
    wbs_id_counter = [1]

    def next_id():
        wid = f"WBS-{wbs_id_counter[0]:04d}"
        wbs_id_counter[0] += 1
        return wid

    # Level 1 — project root
    root_id = next_id()
    rows.append({
        "wbs_id":          root_id,
        "proj_id":         "PROJ-001",
        "obs_id":          "",
        "seq_num":         "1",
        "proj_node_flag":  "Y",
        "sum_data_flag":   "Y",
        "status_code":     "WS_Active",
        "wbs_short_name":  "ALNARJIS-P2",
        "wbs_name":        "Al-Narjis Water Network - Phase 2",
        "parent_wbs_id":   "",
    })

    for z_idx, zone in enumerate(ZONES, 1):
        zone_id = next_id()
        zone_code = f"Z{z_idx:02d}"
        rows.append({
            "wbs_id":         zone_id,
            "proj_id":        "PROJ-001",
            "obs_id":         "",
            "seq_num":        str(z_idx + 1),
            "proj_node_flag": "N",
            "sum_data_flag":  "Y",
            "status_code":    "WS_Active",
            "wbs_short_name": zone_code,
            "wbs_name":       zone,
            "parent_wbs_id":  root_id,
        })

        for d_idx, disc in enumerate(DISCIPLINES, 1):
            disc_id = next_id()
            rows.append({
                "wbs_id":         disc_id,
                "proj_id":        "PROJ-001",
                "obs_id":         "",
                "seq_num":        str(d_idx),
                "proj_node_flag": "N",
                "sum_data_flag":  "N",
                "status_code":    "WS_Active",
                "wbs_short_name": f"{zone_code}-{disc[:3].upper()}",
                "wbs_name":       f"{zone} - {disc}",
                "parent_wbs_id":  zone_id,
            })

    return rows


def _task_rows(wbs_rows: list[dict]) -> list[dict]:
    """Generate ~100 activities distributed across zones and disciplines."""
    random.seed(42)
    rows = []
    task_counter = [1]

    # Build wbs lookup: (zone, disc) -> wbs_id
    wbs_lookup: dict[tuple, str] = {}
    for w in wbs_rows:
        name = w["wbs_name"]
        for zone in ZONES:
            for disc in DISCIPLINES:
                if name == f"{zone} - {disc}":
                    wbs_lookup[(zone, disc)] = w["wbs_id"]

    # Calculate time spans per zone/discipline
    total_days = (PROJECT_END - PROJECT_START).days

    for zone in ZONES:
        zone_idx = ZONES.index(zone)
        planned_pct, actual_pct = ZONE_PROFILES[zone]

        # Zone starts staggered
        zone_offset_days = zone_idx * 15
        zone_start = _add_days(PROJECT_START, zone_offset_days)
        zone_duration_days = total_days - zone_offset_days - 10

        for disc in DISCIPLINES:
            wbs_id = wbs_lookup.get((zone, disc), "")
            disc_seq = DISC_SEQUENCE[disc]

            # Each discipline takes roughly 1/5 of zone duration, in sequence
            disc_offset = (zone_duration_days / len(DISCIPLINES)) * disc_seq
            disc_duration = zone_duration_days / len(DISCIPLINES)

            disc_start = _add_days(zone_start, disc_offset)
            disc_end   = _add_days(disc_start, disc_duration)

            # 2-3 activities per discipline per zone
            n_tasks = random.choice([2, 2, 3])
            for t in range(n_tasks):
                tid = f"TASK-{task_counter[0]:04d}"
                task_counter[0] += 1

                z_num = ZONES.index(zone) + 1
                task_code = f"Z{z_num:02d}-{disc[:3].upper()}-{t+1:03d}"
                task_name = f"{zone} - {disc} - Activity {t+1}"

                # Stagger within discipline
                t_offset_days = (disc_duration / n_tasks) * t
                t_start = _add_days(disc_start, t_offset_days)
                t_end   = _add_days(t_start, disc_duration / n_tasks)
                t_dur   = max(1, int((t_end - t_start).days))

                # Determine status
                if planned_pct == 0 and actual_pct == 0:
                    status = "TK_NotStart"
                    act_start = None
                    act_end   = None
                    phys_pct  = 0.0
                elif actual_pct >= 0.99:
                    status = "TK_Complete"
                    act_start = _add_days(t_start, random.uniform(-2, 2))
                    act_end   = _add_days(t_end, random.uniform(-3, 3))
                    phys_pct  = 1.0
                else:
                    # Activity level progress based on zone profile
                    # Add noise per activity
                    noise = random.uniform(-0.15, 0.10)
                    act_p = max(0.0, min(1.0, actual_pct + noise))

                    if t_start > DATA_DATE:
                        status  = "TK_NotStart"
                        act_start = None
                        act_end   = None
                        phys_pct  = 0.0
                    elif act_p >= 1.0:
                        status    = "TK_Complete"
                        act_start = _add_days(t_start, random.uniform(-1, 1))
                        act_end   = _add_days(t_end, random.uniform(-2, 2))
                        phys_pct  = 1.0
                    elif act_p <= 0.0:
                        status    = "TK_NotStart"
                        act_start = None
                        act_end   = None
                        phys_pct  = 0.0
                    else:
                        status    = "TK_Active"
                        act_start = _add_days(t_start, random.uniform(-2, 3))
                        act_end   = None
                        phys_pct  = round(act_p, 4)

                remaining = t_dur * (1.0 - phys_pct)

                rows.append({
                    "task_id":            tid,
                    "proj_id":            "PROJ-001",
                    "wbs_id":             wbs_id,
                    "task_code":          task_code,
                    "task_name":          task_name,
                    "task_type":          "TT_Task",
                    "status_code":        status,
                    "target_start_date":  _fmt(t_start),
                    "target_end_date":    _fmt(t_end),
                    "act_start_date":     _fmt(act_start),
                    "act_end_date":       _fmt(act_end),
                    "target_drtn_hr_cnt": str(t_dur * 8),
                    "remain_drtn_hr_cnt": str(max(0, int(remaining * 8))),
                    "phys_complete_pct":  str(phys_pct),
                    "total_float_hr_cnt": str(random.randint(-40, 40)),
                })

    return rows


def _taskrsrc_rows(task_rows: list[dict]) -> list[dict]:
    """Assign resources to activities."""
    random.seed(42)
    rows = []
    tr_counter = [1]

    rsrc_map = {r[0]: r for r in RESOURCES}

    for task in task_rows:
        task_code = task["task_code"]
        # Identify discipline from task code prefix
        disc_abbr = task_code.split("-")[1] if "-" in task_code else ""
        disc_abbr_map = {
            "EXC": "Excavation",
            "PIP": "Pipe Laying",
            "BAC": "Backfill",
            "REI": "Reinstatement",
            "TES": "Testing & Commissioning",
        }
        disc = disc_abbr_map.get(disc_abbr, "")
        rsrc_ids = DISC_RESOURCES.get(disc, [])

        if not rsrc_ids:
            continue

        phys_pct = float(task.get("phys_complete_pct", 0))

        for rsrc_id in rsrc_ids:
            rsrc = rsrc_map[rsrc_id]
            unit_price = rsrc[4]

            target_qty = round(random.uniform(50, 500), 2)
            act_qty    = round(target_qty * phys_pct * random.uniform(0.9, 1.1), 2)
            remain_qty = round(max(0, target_qty - act_qty), 2)

            target_cost = round(target_qty * unit_price, 2)
            act_cost    = round(act_qty * unit_price, 2)
            remain_cost = round(remain_qty * unit_price, 2)

            rows.append({
                "taskrsrc_id":   f"TR-{tr_counter[0]:05d}",
                "task_id":       task["task_id"],
                "proj_id":       "PROJ-001",
                "rsrc_id":       rsrc_id,
                "rsrc_type":     rsrc[2],
                "target_qty":    str(target_qty),
                "act_reg_qty":   str(act_qty),
                "act_ot_qty":    "0.0",
                "remain_qty":    str(remain_qty),
                "target_cost":   str(target_cost),
                "act_reg_cost":  str(act_cost),
                "act_ot_cost":   "0.0",
                "remain_cost":   str(remain_cost),
                "cost_per_qty":  str(unit_price),
            })
            tr_counter[0] += 1

    return rows


def _rsrc_rows() -> list[dict]:
    rows = []
    for rsrc_id, rsrc_name, rsrc_type, uom, price in RESOURCES:
        rows.append({
            "rsrc_id":       rsrc_id,
            "rsrc_name":     rsrc_name,
            "rsrc_short_name": rsrc_id,
            "rsrc_type":     rsrc_type,
            "unit_id":       uom,
            "cost_per_qty":  str(price),
            "rsrc_notes":    "",
        })
    return rows


# ── XER writer ────────────────────────────────────────────────────────────────

def _write_table(lines: list[str], table_name: str, rows: list[dict]) -> None:
    if not rows:
        return
    fields = list(rows[0].keys())
    lines.append(f"%T\t{table_name}")
    lines.append("%F\t" + "\t".join(fields))
    for row in rows:
        lines.append("%R\t" + "\t".join(str(row.get(f, "")) for f in fields))


def generate_sample_xer(output_path: str | Path) -> Path:
    """Generate a realistic sample XER and write it to output_path."""
    output_path = Path(output_path)

    wbs   = _wbs_rows()
    tasks = _task_rows(wbs)
    rsrcs = _rsrc_rows()
    trsrc = _taskrsrc_rows(tasks)

    lines: list[str] = [
        "ERMHDR\t19.12\t2026-04-01\tPrimavera\tadmin\t\t\tProject\t\t",
    ]

    _write_table(lines, "PROJECT",  [_project_row()])
    _write_table(lines, "WBS",      wbs)
    _write_table(lines, "TASK",     tasks)
    _write_table(lines, "RSRC",     rsrcs)
    _write_table(lines, "TASKRSRC", trsrc)

    lines.append("%E")

    output_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"[sample_generator] Written {len(tasks)} tasks, "
          f"{len(trsrc)} resource assignments → {output_path}")
    return output_path
