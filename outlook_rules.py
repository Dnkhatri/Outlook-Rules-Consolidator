"""
Outlook Rules Consolidator & Restorer
======================================
Pure Python / COM — no AI, no external API.

How your ruleset works
-----------------------
You have catch-all rules at the end of your list, each named after their
target folder (e.g. rule "Local" moves to folder "Local", rule "Vietnam"
moves to folder "Vietnam"). Some have alias names ("Warehouse" → SS&W,
"Computers" → "Computers and Web", etc.).

When you create new individual sender rules in Outlook, running this
script folds them into the appropriate catch-all's From condition and
deletes the individual rules.

Running it repeatedly is safe — senders already in a catch-all are
never added twice.

Usage
-----
    python outlook_rules.py --dry-run              # safe preview
    python outlook_rules.py --apply                # do it
    python outlook_rules.py --restore backup.json  # restore from backup
    python outlook_rules.py --restore backup.json --dry-run

Requirements
------------
    pip install pywin32
"""

import argparse
import json
import logging
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

import win32com.client

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

BACKUP_DIR = Path(".")

# ---------------------------------------------------------------------------
# Catch-all name aliases
# A catch-all whose rule name differs from its folder name goes here.
# Add new aliases if you create catch-alls with non-matching names.
# ---------------------------------------------------------------------------
CATCHALL_ALIASES: Dict[str, str] = {
    "Warehouse":    "SS&W",
    "indian":       "Indian",
    "Bangla Nepal": "Bangla&Nepal",
    "Computers":    "Computers and Web",
}


# ---------------------------------------------------------------------------
# COM helpers
# ---------------------------------------------------------------------------

def _save(com_rules, context: str = "") -> None:
    try:
        com_rules.Save()
        log.info("Saved%s.", f" ({context})" if context else "")
    except Exception as exc:
        log.error("Save failed%s: %s", f" ({context})" if context else "", exc)


def _find_folder(session, folder_name: str):
    """Recursively find a mail folder by name in the default store."""
    def _search(folder):
        if folder.Name == folder_name:
            return folder
        try:
            for i in range(1, folder.Folders.Count + 1):
                hit = _search(folder.Folders.Item(i))
                if hit is not None:
                    return hit
        except Exception:
            pass
        return None
    try:
        return _search(session.DefaultStore.GetRootFolder())
    except Exception as exc:
        log.debug("Folder search error '%s': %s", folder_name, exc)
        return None


def _get_from_recipient_names(com_rule) -> Set[str]:
    """Return the set of names already in a rule's From condition (lowercased)."""
    names: Set[str] = set()
    try:
        from_cond = com_rule.Conditions.From
        if from_cond.Enabled:
            for r in from_cond.Recipients:
                names.add(r.Name.strip().lower())
    except Exception:
        pass
    return names


# ---------------------------------------------------------------------------
# Rule extraction
# ---------------------------------------------------------------------------

def extract_rule(rule, com_index: int) -> Dict:
    info: Dict[str, Any] = {
        "com_index":       com_index,
        "name":            rule.Name,
        "enabled":         rule.Enabled,
        "execution_order": rule.ExecutionOrder,
        "actions":         {},
        "_com_rule":       rule,
        "_folder_obj":     None,
    }
    try:
        act = rule.Actions
        try:
            if act.MoveToFolder.Enabled:
                info["actions"]["move_to_folder"] = act.MoveToFolder.Folder.Name
                info["_folder_obj"] = act.MoveToFolder.Folder
        except Exception:
            pass
        try:
            if act.Delete.Enabled:
                info["actions"]["delete"] = True
        except Exception:
            pass
        try:
            if act.MarkAsRead.Enabled:
                info["actions"]["mark_as_read"] = True
        except Exception:
            pass
        try:
            if act.FlagMessage.Enabled:
                info["actions"]["flag_message"] = True
        except Exception:
            pass
        try:
            if act.ForwardTo.Enabled:
                info["actions"]["forward_to"] = [r.Address for r in act.ForwardTo.Recipients]
        except Exception:
            pass
    except Exception as exc:
        log.debug("Action read error [%d]: %s", com_index, exc)
    return info


def extract_all(com_rules) -> List[Dict]:
    rules = []
    for i in range(1, com_rules.Count + 1):
        try:
            rules.append(extract_rule(com_rules.Item(i), i))
        except Exception as exc:
            log.warning("Skipping rule %d: %s", i, exc)
    log.info("Extracted %d rules.", len(rules))
    return rules


# ---------------------------------------------------------------------------
# Auto-detect catch-all rules
# ---------------------------------------------------------------------------

def detect_catchalls(rules: List[Dict]) -> Dict[str, Dict]:
    """
    Auto-detect catch-all rules. A rule is a catch-all when:
      - Its name matches its target folder name  (e.g. "Vietnam" → "Vietnam")
      - OR its name is a known alias             (e.g. "Warehouse" → "SS&W")

    Returns {folder_name: rule_data}.
    If multiple rules qualify for the same folder, the last one wins
    (matches the pattern of your ruleset where catch-alls are at the end).
    """
    catchalls: Dict[str, Dict] = {}
    for r in rules:
        name   = r["name"]
        folder = r["actions"].get("move_to_folder")

        if not folder:
            continue

        # Direct match: rule name == folder name
        if name == folder:
            catchalls[folder] = r
            continue

        # Alias match
        if name in CATCHALL_ALIASES and CATCHALL_ALIASES[name] == folder:
            catchalls[folder] = r

    return catchalls


def is_sender_rule(r: Dict, catchall_indices: Set[int]) -> bool:
    """Individual sender rules: move-only, not a catch-all, not complex."""
    return (
        r["com_index"] not in catchall_indices
        and "move_to_folder" in r["actions"]
        and len(r["actions"]) == 1      # only MoveToFolder, nothing else
    )


# ---------------------------------------------------------------------------
# Consolidation
# ---------------------------------------------------------------------------

def build_merge_plan(rules: List[Dict]) -> Dict:
    """
    Work out what needs merging without touching Outlook.

    Returns a dict:
      {
        folder_name: {
          "catchall": rule_data,
          "to_add":   [name, ...],   # new senders not yet in catch-all
          "to_delete": [com_index, ...]
        },
        ...
        "_no_catchall": [rule_data, ...]   # sender rules with no catch-all
        "_already_merged": [rule_data, ...] # sender rules already in catch-all
      }
    """
    catchalls    = detect_catchalls(rules)
    catchall_idx = {v["com_index"] for v in catchalls.values()}
    sender_rules = [r for r in rules if is_sender_rule(r, catchall_idx)]

    plan: Dict[str, Any] = {"_no_catchall": [], "_already_merged": []}

    for folder, ca in catchalls.items():
        plan[folder] = {
            "catchall":  ca,
            "to_add":    [],
            "to_delete": [],
        }

    # Get existing recipients for each catch-all (to avoid duplicates)
    existing: Dict[str, Set[str]] = {}
    for folder, ca in catchalls.items():
        existing[folder] = _get_from_recipient_names(ca["_com_rule"])

    for r in sender_rules:
        folder = r["actions"]["move_to_folder"]
        if folder not in catchalls:
            plan["_no_catchall"].append(r)
            continue

        sender_key = r["name"].strip().lower()
        if sender_key in existing.get(folder, set()):
            plan["_already_merged"].append(r)
            # Still delete — it's redundant even if already in catch-all
            plan[folder]["to_delete"].append(r["com_index"])
        else:
            plan[folder]["to_add"].append(r["name"])
            plan[folder]["to_delete"].append(r["com_index"])

    return plan


def print_preview(rules: List[Dict]) -> None:
    plan         = build_merge_plan(rules)
    catchalls    = detect_catchalls(rules)
    catchall_idx = {v["com_index"] for v in catchalls.values()}
    sender_rules = [r for r in rules if is_sender_rule(r, catchall_idx)]

    total_to_delete  = sum(len(plan[f]["to_delete"]) for f in catchalls)
    total_new        = sum(len(plan[f]["to_add"]) for f in catchalls)
    already_merged   = len(plan["_already_merged"])
    no_catchall      = plan["_no_catchall"]

    print(f"\n{'='*64}\n📦  CONSOLIDATION PREVIEW\n{'='*64}")
    print(f"  Rules now                 : {len(rules)}")
    print(f"  Catch-alls detected       : {len(catchalls)}")
    print(f"  Individual rules to merge : {len(sender_rules)}")
    print(f"    → New senders to add    : {total_new}")
    print(f"    → Already in catch-all  : {already_merged}  (will just be deleted)")
    print(f"  Rules after               : {len(rules) - total_to_delete}")
    print()

    for folder in sorted(catchalls):
        entry = plan.get(folder, {})
        to_add = entry.get("to_add", [])
        to_del = entry.get("to_delete", [])
        ca     = catchalls[folder]
        if to_add or to_del:
            print(f"  [{folder}]  catch-all='{ca['name']}'  "
                  f"+{len(to_add)} new, -{len(to_del)} individual rules")
            for s in to_add:
                print(f"      + {s}")

    if no_catchall:
        print(f"\n  ⚠️  {len(no_catchall)} rules have no matching catch-all "
              f"(left untouched):")
        for r in no_catchall:
            print(f"      • {r['name']}  →  {r['actions'].get('move_to_folder','?')}")
        print("    To fix: create a catch-all rule named after the folder,")
        print("    then re-run this script.")

    print("="*64 + "\n")


def apply_consolidation(rules: List[Dict], com_rules, dry_run: bool) -> None:
    plan         = build_merge_plan(rules)
    catchalls    = detect_catchalls(rules)
    catchall_idx = {v["com_index"] for v in catchalls.values()}
    sender_rules = [r for r in rules if is_sender_rule(r, catchall_idx)]

    if not sender_rules:
        log.info("Nothing to consolidate — all sender rules already merged or absent.")
        return

    total_new = sum(len(plan[f]["to_add"]) for f in catchalls)
    total_del = sum(len(plan[f]["to_delete"]) for f in catchalls)

    if dry_run:
        print_preview(rules)
        log.info("Dry run — no changes made.")
        return

    # --- Step 1: add new senders to each catch-all ---
    all_to_delete: List[int] = []

    for folder in sorted(catchalls):
        entry  = plan.get(folder, {})
        to_add = entry.get("to_add", [])
        to_del = entry.get("to_delete", [])

        all_to_delete.extend(to_del)

        if not to_add:
            continue

        ca       = catchalls[folder]
        com_rule = ca["_com_rule"]
        try:
            from_cond = com_rule.Conditions.From
            from_cond.Enabled = True
            added = 0
            for name in to_add:
                try:
                    from_cond.Recipients.Add(name)
                    added += 1
                except Exception as e:
                    log.debug("Could not add '%s' to '%s': %s", name, folder, e)
            try:
                from_cond.Recipients.ResolveAll()
            except Exception:
                pass
            log.info("Added %d sender(s) to '%s' catch-all.", added, folder)
        except Exception as exc:
            log.error("Failed updating catch-all for '%s': %s", folder, exc)

    _save(com_rules, "after updating catch-all conditions")

    # --- Step 2: delete individual rules (highest index first) ---
    if all_to_delete:
        log.info("Deleting %d individual sender rules...", len(all_to_delete))
        for idx in sorted(set(all_to_delete), reverse=True):
            try:
                com_rules.Remove(idx)
            except Exception as exc:
                log.error("Delete [%d] failed: %s", idx, exc)
        _save(com_rules, "after deleting individual rules")

    no_catchall = plan["_no_catchall"]
    if no_catchall:
        log.warning(
            "%d rule(s) had no matching catch-all and were left untouched: %s",
            len(no_catchall),
            [r["name"] for r in no_catchall],
        )

    log.info(
        "Done. Added %d senders, deleted %d rules. Outlook now has %d rules.",
        total_new, total_del, com_rules.Count,
    )


# ---------------------------------------------------------------------------
# Restore from backup
# ---------------------------------------------------------------------------

def apply_restore(backup_path: Path, com_rules, session, dry_run: bool) -> None:
    """Wipe all current rules and recreate from backup JSON."""
    if not backup_path.exists():
        log.error("Backup not found: %s", backup_path.resolve())
        sys.exit(1)
    try:
        backed_up: List[Dict] = json.loads(backup_path.read_text(encoding="utf-8"))
    except Exception as exc:
        log.error("Could not read backup: %s", exc)
        sys.exit(1)

    backed_up.sort(key=lambda r: r.get("execution_order", 9999))
    current_count = com_rules.Count
    move_rules    = [r for r in backed_up if "move_to_folder" in r.get("actions", {})]
    other_rules   = [r for r in backed_up if "move_to_folder" not in r.get("actions", {})]

    if dry_run:
        by_folder: Dict[str, List[str]] = defaultdict(list)
        for r in move_rules:
            by_folder[r["actions"]["move_to_folder"]].append(r["name"])
        print(f"\n{'='*64}\n🔍  DRY RUN — restore from {backup_path.name}\n{'='*64}")
        print(f"  Would DELETE  : {current_count} current rules")
        print(f"  Would RESTORE : {len(backed_up)} rules\n")
        for folder in sorted(by_folder):
            print(f"  [{folder}]  {len(by_folder[folder])} rules")
        if other_rules:
            print(f"\n  Non-move rules (disabled stubs):")
            for r in other_rules:
                print(f"    • {r['name']}  {list(r.get('actions', {}).keys())}")
        print("="*64 + "\n")
        return

    confirm = input(
        f"\n⚠️  DELETE all {current_count} current rules and restore "
        f"{len(backed_up)} from backup.\nType CONFIRM: "
    ).strip()
    if confirm != "CONFIRM":
        log.info("Cancelled.")
        return

    log.info("Wiping %d current rules...", current_count)
    for idx in range(current_count, 0, -1):
        try:
            com_rules.Remove(idx)
        except Exception as exc:
            log.error("Delete [%d] failed: %s", idx, exc)
    _save(com_rules, "after wipe")

    stubs: List[str] = []
    for r in backed_up:
        name    = r["name"]
        act     = r.get("actions", {})
        enabled = r.get("enabled", True)
        try:
            new_rule = com_rules.Create(name, 0)
            if "move_to_folder" in act:
                folder_name = act["move_to_folder"]
                folder_obj  = _find_folder(session, folder_name)
                if folder_obj:
                    from_cond = new_rule.Conditions.From
                    from_cond.Enabled = True
                    try:
                        from_cond.Recipients.Add(name)
                        from_cond.Recipients.ResolveAll()
                    except Exception:
                        pass
                    move = new_rule.Actions.MoveToFolder
                    move.Enabled = True
                    move.Folder  = folder_obj
                    new_rule.Enabled = enabled
                else:
                    new_rule.Enabled = False
                    stubs.append(f"{name}  (folder '{folder_name}' not found)")
            else:
                new_rule.Enabled = False
                stubs.append(f"{name}  ({list(act.keys())})")
        except Exception as exc:
            log.error("Failed to restore '%s': %s", name, exc)

    _save(com_rules, "after restore")
    log.info("Restore complete. Outlook now has %d rules.", com_rules.Count)
    if stubs:
        print(f"\n⚠️  {len(stubs)} rules need manual config in Outlook:")
        for s in stubs:
            print(f"   • {s}")


# ---------------------------------------------------------------------------
# Backup & summary
# ---------------------------------------------------------------------------

def backup(rules: List[Dict]) -> Path:
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = BACKUP_DIR / f"outlook_rules_backup_{ts}.json"
    clean = [{k: v for k, v in r.items() if not k.startswith("_")} for r in rules]
    path.write_text(json.dumps(clean, indent=2, default=str), encoding="utf-8")
    log.info("Backup → %s", path.resolve())
    return path


def print_summary(rules: List[Dict]) -> None:
    catchalls    = detect_catchalls(rules)
    catchall_idx = {v["com_index"] for v in catchalls.values()}
    senders      = [r for r in rules if is_sender_rule(r, catchall_idx)]
    other        = [r for r in rules if r["com_index"] not in catchall_idx
                    and not is_sender_rule(r, catchall_idx)]
    print(f"\n{'='*64}\n📋  RULES SUMMARY\n{'='*64}")
    print(f"  Total rules              : {len(rules)}")
    print(f"  Catch-alls detected      : {len(catchalls)}")
    print(f"    {sorted(catchalls.keys())}")
    print(f"  Individual sender rules  : {len(senders)}  ← will be merged")
    print(f"  Other rules              : {len(other)}")
    print(f"  After consolidation      : {len(rules) - len(senders)}")
    print("="*64 + "\n")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Merge individual sender rules into catch-all rules. Safe to re-run."
    )
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", action="store_true",
                      help="Preview changes without modifying anything (safe)")
    mode.add_argument("--apply",   action="store_true",
                      help="Apply changes without prompting")
    p.add_argument("--restore",    metavar="BACKUP_JSON",
                   help="Restore all rules from a backup JSON")
    p.add_argument("--no-backup",  action="store_true",
                   help="Skip creating a JSON backup before changes")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    log.info("Outlook Rules Consolidator")

    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        session   = outlook.Session
        com_rules = session.DefaultStore.GetRules()
        log.info("Connected. Outlook has %d rules.", com_rules.Count)
    except Exception as exc:
        log.error("Cannot connect to Outlook: %s", exc)
        sys.exit(1)

    if args.restore:
        apply_restore(Path(args.restore), com_rules, session, dry_run=args.dry_run)
        return

    rules = extract_all(com_rules)
    if not rules:
        log.info("No rules found.")
        sys.exit(0)

    print_summary(rules)

    if not args.no_backup:
        backup(rules)

    if args.apply:
        dry = False
    elif args.dry_run:
        dry = True
    else:
        print_preview(rules)
        print("⚠️   This will modify your Outlook rules.")
        choice = input("Proceed? [dry-run / yes / no]: ").strip().lower()
        if choice in ("yes", "y"):
            if input("Type CONFIRM: ").strip() != "CONFIRM":
                log.info("Cancelled.")
                sys.exit(0)
            dry = False
        elif choice in ("dry-run", "d"):
            dry = True
        else:
            log.info("No changes made.")
            sys.exit(0)

    apply_consolidation(rules, com_rules, dry_run=dry)

    if not dry:
        log.info("Verify in Outlook → File → Manage Rules & Alerts.")


if __name__ == "__main__":
    main()
