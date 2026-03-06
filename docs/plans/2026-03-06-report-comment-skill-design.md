# Design: Report Comment Skill

**Date:** 2026-03-06

## Overview

A Claude Code skill that generates structured template comments for the comment bank application. The user provides a Subject and Topic, Claude generates High/Medium/Low comments using standard template variables, refines interactively, then saves the result to a markdown file.

## Inputs

- **Subject** — e.g. `Computing`
- **Topic** — e.g. `Theoretic Knowledge` (also used as the comment set name)

## Template Variables

Available for use in all comments:

`<Name>`, `<he/she>`, `<his/her>`, `<him/her>`, `<Subject>`, `<Year>`, `<EoYLevel>`, `<TargetLevel>`

## Comment Structure

- **3 levels:** High, Medium, Low
- **1–2 sentences per level**
- **Tone:** professional and supportive — no negative language at any level

## Workflow

1. User invokes skill with Subject and Topic
2. Claude generates a draft set of 3 comments
3. Claude presents the draft and asks: *"Are you happy with these, or would you like any level adjusted?"*
4. User can request changes per level or approve
5. Refinement loop continues until user approves
6. Claude displays the final approved set as a markdown table
7. Claude writes the file to `comments/<Subject>/<Topic>.md`

## Output File Format

Path: `comments/<Subject>/<Topic>.md`

Example: `comments/Computing/Theoretic Knowledge.md`

```markdown
# <Subject> — <Topic>

## High
<comment>

## Medium
<comment>

## Low
<comment>
```

## Example

**Input:** Subject = Computing, Topic = Theoretic Knowledge

**Output file:** `comments/Computing/Theoretic Knowledge.md`

| Level | Comment |
|-------|---------|
| High | `<Name> has demonstrated an excellent understanding of theoretical concepts in <Subject> this year. <He/She> engages with complex ideas confidently and is working well above <his/her> target of <TargetLevel>.` |
| Medium | `<Name> has shown a solid understanding of theoretical concepts in <Subject> and is making good progress. <He/She> is on track to meet <his/her> target of <TargetLevel> by the end of <Year>.` |
| Low | `<Name> is developing <his/her> understanding of theoretical concepts in <Subject> and is working towards <his/her> target of <TargetLevel>. With continued support, <he/she> is making progress this year.` |
