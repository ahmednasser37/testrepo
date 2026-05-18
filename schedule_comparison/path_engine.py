"""
Longest path / Bottleneck engine.
Identifies the activity with the greatest Remaining Path Length (RPL)
at the data date — the current network bottleneck.
"""
from __future__ import annotations
import pandas as pd


def compute_longest_path(updated_tables: dict, data_date) -> dict:
    """
    Returns:
        {
          "bottleneck": {
              "task_code": str, "task_name": str, "wbs_name": str,
              "rpl_days": float, "planned_finish": str,
              "total_float_days": float, "status": str
          } | None,
          "driving_chain": [
              {"task_code", "task_name", "planned_finish", "rpl_days", "depth": int}, ...
          ]  (up to 15 activities, driving predecessor first)
        }
    """
    task_df = updated_tables.get("TASK", pd.DataFrame())
    pred_df = updated_tables.get("TASKPRED", pd.DataFrame())
    wbs_df  = updated_tables.get("PROJWBS", updated_tables.get("WBS", pd.DataFrame()))

    if task_df.empty:
        return {"bottleneck": None, "driving_chain": []}

    dd = pd.Timestamp(data_date) if data_date else pd.Timestamp.now()

    # WBS name lookup
    wbs_names: dict[str, str] = {}
    if not wbs_df.empty and "wbs_id" in wbs_df.columns:
        wbs_names = dict(zip(wbs_df["wbs_id"].astype(str), wbs_df["wbs_name"].astype(str)))

    # Build task index by task_id
    task_idx: dict[str, dict] = {}
    for _, row in task_df.iterrows():
        tid = str(row.get("task_id", ""))
        task_idx[tid] = {
            "task_id":      tid,
            "task_code":    str(row.get("task_code", "")),
            "task_name":    str(row.get("task_name", "")),
            "wbs_id":       str(row.get("wbs_id", "")),
            "status_code":  str(row.get("status_code", "")),
            "target_end_date": pd.to_datetime(row.get("target_end_date"), errors="coerce"),
            "total_float_hr_cnt": float(row.get("total_float_hr_cnt", 0) or 0),
        }

    # Build successor index: task_id → [(succ_task_id, lag_days)]
    succ_idx: dict[str, list] = {}
    if not pred_df.empty:
        for _, rel in pred_df.iterrows():
            ptype = str(rel.get("pred_type", ""))
            if ptype not in ("PR_FS", "FS", "pr_fs"):
                continue  # only FS for longest path
            pid  = str(rel.get("pred_task_id", ""))
            sid  = str(rel.get("task_id", ""))
            lag  = float(rel.get("lag_hr_cnt", 0) or 0) / 8.0
            succ_idx.setdefault(pid, []).append((sid, lag))

    # Memoized DFS to compute RPL for each task
    memo: dict[str, float] = {}
    visited: set[str] = set()

    def rpl(tid: str) -> float:
        if tid in memo:
            return memo[tid]
        if tid in visited:
            memo[tid] = 0.0
            return 0.0
        visited.add(tid)

        task = task_idx.get(tid)
        if task is None:
            memo[tid] = 0.0
            return 0.0

        status = task["status_code"]
        if status == "TK_Complete":
            memo[tid] = 0.0
            visited.discard(tid)
            return 0.0

        pf = task["target_end_date"]
        own_days = max(0.0, (pf - dd).total_seconds() / 86400.0) if pd.notna(pf) else 0.0

        best_succ = 0.0
        for (sid, lag) in succ_idx.get(tid, []):
            candidate = lag + rpl(sid)
            if candidate > best_succ:
                best_succ = candidate

        result = own_days + best_succ
        memo[tid] = result
        visited.discard(tid)
        return result

    # Compute RPL for all non-complete tasks
    best_rpl   = -1.0
    bottleneck_id = None
    for tid, task in task_idx.items():
        if task["status_code"] == "TK_Complete":
            continue
        r = rpl(tid)
        if r > best_rpl:
            best_rpl = r
            bottleneck_id = tid

    if bottleneck_id is None:
        return {"bottleneck": None, "driving_chain": []}

    bt = task_idx[bottleneck_id]
    pf = bt["target_end_date"]
    bottleneck = {
        "task_code":        bt["task_code"],
        "task_name":        bt["task_name"],
        "wbs_name":         wbs_names.get(bt["wbs_id"], ""),
        "rpl_days":         round(best_rpl, 1),
        "planned_finish":   pf.strftime("%Y-%m-%d") if pd.notna(pf) else "",
        "total_float_days": round(bt["total_float_hr_cnt"] / 8.0, 1),
        "status":           bt["status_code"],
    }

    # Build driving chain (walk back from bottleneck via its own predecessors that maximise RPL)
    # First build predecessor index
    pred_idx: dict[str, list] = {}  # task_id → [(pred_task_id, lag_days)]
    if not pred_df.empty:
        for _, rel in pred_df.iterrows():
            ptype = str(rel.get("pred_type", ""))
            if ptype not in ("PR_FS", "FS", "pr_fs"):
                continue
            pid = str(rel.get("pred_task_id", ""))
            sid = str(rel.get("task_id", ""))
            lag = float(rel.get("lag_hr_cnt", 0) or 0) / 8.0
            pred_idx.setdefault(sid, []).append((pid, lag))

    chain = []
    chain_visited: set[str] = set()

    def build_chain(tid: str, depth: int):
        if depth > 14 or tid in chain_visited:
            return
        chain_visited.add(tid)
        task = task_idx.get(tid)
        if task is None:
            return
        pf = task["target_end_date"]
        chain.append({
            "task_code":      task["task_code"],
            "task_name":      task["task_name"],
            "planned_finish": pf.strftime("%Y-%m-%d") if pd.notna(pf) else "",
            "rpl_days":       round(memo.get(tid, 0.0), 1),
            "depth":          depth,
        })
        # Walk to driving predecessor (pred with highest RPL contribution)
        preds = pred_idx.get(tid, [])
        if not preds:
            return
        best_pred_id = max(preds, key=lambda x: rpl(x[0]) + x[1])[0]
        build_chain(best_pred_id, depth + 1)

    build_chain(bottleneck_id, 0)

    return {"bottleneck": bottleneck, "driving_chain": chain}
