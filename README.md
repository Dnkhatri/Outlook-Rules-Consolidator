# Outlook Rule Consolidator

## What It Does

Over time, Outlook rules build up. You might have 20 separate rules all doing the same thing — moving emails from different Vietnam suppliers into your Vietnam folder. Another 15 rules all moving emails into your Shippers folder. And so on. Before long you have hundreds of rules that are hard to manage.

This script cleans that up. It looks at all your rules, groups together every rule that moves emails into the same folder, and merges them into one single rule per folder. So instead of 20 separate Vietnam rules, you end up with just one Vietnam rule that covers all 20 suppliers. Instead of 15 Shippers rules, one Shippers rule. Your emails still land in exactly the right folders — there are just far fewer rules doing it.

The only thing you need to do to make this work is follow one simple naming convention when creating rules, explained in the next section.

---

## Creating Rules the Right Way

Every time you create a new rule in Outlook:

1. Go to **File → Manage Rules & Alerts → New Rule**
2. Set the condition — who the email is from
3. Set the action — which folder to move it to
4. **Name the rule exactly the same as the folder**

For example, if the email should go to your **Vietnam** folder, name the rule **Vietnam**.

That's the only convention you need to follow whne create the main rule in the future when ever you create a rule that moves the email to that folder the script would consolidate it to the main rule. The script handles everything else.

---

## Running the Script

### First time — preview before doing anything

```
python outlook_rules.py --dry-run
```

This shows you exactly what would be merged and deleted without touching anything in Outlook. Always run this first to make sure everything looks right.

### Apply the consolidation

```
python outlook_rules.py --apply
```

The script will back up all your current rules to a file (e.g. `outlook_rules_backup_20260519_144519.json`) before making any changes. Keep this file safe.

### In the future — after adding new rules

Every time your rules list has grown and needs tidying up again, just run:

```
python outlook_rules.py --apply
```

It is safe to run as many times as you like. It will never add the same contact twice.

---

## Restoring Your Rules

If anything goes wrong and you want to go back to exactly how things were before:

```
python outlook_rules.py --restore outlook_rules_backup_20260519_144519.json
```

Replace the filename with the actual backup file that was created on your machine. Type `CONFIRM` when prompted.

---

## Requirements

- Windows PC with Outlook open
- Python 3 — download from https://www.python.org/downloads/
- Run once to install the required library:

```
pip install pywin32
```
