# Report Comment Skill Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Create a Claude Code skill that interactively generates High/Medium/Low school report comment templates for a comment bank, then saves them to a markdown file.

**Architecture:** A single SKILL.md file installed at `~/.claude/skills/report-comment/SKILL.md`. The skill instructs Claude to collect Subject and Topic from the user, generate 3 levelled comments using template variables, run an approval/refinement loop, then display a final table and write the output file.

**Tech Stack:** Markdown skill file, Write tool for file output, `comments/<Subject>/<Topic>.md` output path relative to the comment-bank project root.

---

### Task 1: Create skill directory and SKILL.md

**Files:**
- Create: `~/.claude/skills/report-comment/SKILL.md`

**Step 1: Create the skill directory**

```bash
mkdir -p ~/.claude/skills/report-comment
```

**Step 2: Write the SKILL.md file**

Create `~/.claude/skills/report-comment/SKILL.md` with the following content exactly:

```markdown
---
name: report-comment
description: Use when a user wants to create school report comment templates for a comment bank, given a subject and topic. Generates High, Medium, and Low level comments using standard template variables, refines interactively, then saves to a markdown file.
---

# Report Comment Generator

## Overview

Generate a set of three levelled school report comment templates (High, Medium, Low) for a given subject and topic. Comments use standard template variables, are professional and supportive in tone, and are saved to a markdown file in the comment bank.

## Template Variables

All comments may use any of these variables:

| Variable | Meaning |
|----------|---------|
| `<Name>` | Pupil's name |
| `<he/she>` | Gender pronoun (subject) |
| `<his/her>` | Gender pronoun (possessive) |
| `<him/her>` | Gender pronoun (object) |
| `<Subject>` | Subject name |
| `<Year>` | Academic year |
| `<EoYLevel>` | End of year level achieved |
| `<TargetLevel>` | Target level |

## Comment Rules

- **3 levels:** High, Medium, Low
- **1–2 sentences per level**
- **Tone:** professional and supportive — no negative language, even at Low level
- Low comments should frame development positively (e.g. "is developing", "is working towards", "with support")

## Workflow

### Step 1 — Collect inputs

Ask the user for:
- **Subject** (e.g. Computing)
- **Topic** (e.g. Theoretic Knowledge)

If both are provided when the skill is invoked (e.g. `/report-comment Computing, Theoretic Knowledge`), skip asking.

### Step 2 — Generate draft comments

Generate 3 comments using the rules above. Present them clearly:

**Draft Comments: [Subject] — [Topic]**

| Level | Comment |
|-------|---------|
| High | ... |
| Medium | ... |
| Low | ... |

Then ask:
> "Are you happy with these comments, or would you like any level adjusted?"

### Step 3 — Refinement loop

- If the user requests changes to one or more levels, regenerate only those levels and show the updated table
- Ask again for approval
- Repeat until the user approves all three

### Step 4 — Display final table

Once approved, display:

**Final Comments: [Subject] — [Topic]**

| Level | Comment |
|-------|---------|
| High | ... |
| Medium | ... |
| Low | ... |

### Step 5 — Write to file

Write the approved comments to:

```
comments/<Subject>/<Topic>.md
```

relative to the comment-bank project root (`/Users/leroysalih/nodejs/comment-bank/`).

Create the subject directory if it does not exist.

**File format:**

```markdown
# <Subject> — <Topic>

## High
<high comment>

## Medium
<medium comment>

## Low
<low comment>
```

Confirm to the user: "Saved to `comments/<Subject>/<Topic>.md`."

## Example

**Input:** Subject = Computing, Topic = Theoretic Knowledge

**Output file:** `comments/Computing/Theoretic Knowledge.md`

```markdown
# Computing — Theoretic Knowledge

## High
<Name> has demonstrated an excellent understanding of theoretical concepts in <Subject> this year. <He/She> engages with complex ideas confidently and consistently performs above <his/her> target of <TargetLevel>.

## Medium
<Name> has shown a solid understanding of theoretical concepts in <Subject> and is making good progress. <He/She> is on track to meet <his/her> target of <TargetLevel> by the end of <Year>.

## Low
<Name> is developing <his/her> understanding of theoretical concepts in <Subject> and is working towards <his/her> target of <TargetLevel>. With continued support, <he/she> is making encouraging progress this year.
```
```

**Step 3: Verify the file exists**

```bash
cat ~/.claude/skills/report-comment/SKILL.md
```

Expected: Full SKILL.md content printed to terminal.

**Step 4: Commit**

```bash
cd /Users/leroysalih/nodejs/comment-bank
git add docs/plans/2026-03-06-report-comment-skill.md
git commit -m "docs: add report-comment skill implementation plan"
```

---

### Task 2: Verify skill is discoverable

**Step 1: Check skill appears in Claude Code**

In Claude Code, run:
```
/report-comment
```

Expected: Claude loads the skill and asks for Subject and Topic.

**Step 2: Test with a sample invocation**

Invoke with:
```
/report-comment Computing, Theoretic Knowledge
```

Expected: Claude generates a High/Medium/Low table and enters the refinement loop.

**Step 3: Approve and verify file output**

Approve the comments and verify the file is written:

```bash
cat "/Users/leroysalih/nodejs/comment-bank/comments/Computing/Theoretic Knowledge.md"
```

Expected: Markdown file with `# Computing — Theoretic Knowledge` header and three level sections.
